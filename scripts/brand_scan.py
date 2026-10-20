"""Brand-name scan of a SAM.gov notice's full description (and, optionally, its public attachments).

Why: the SGS search result only carries the first ~200 characters of the description, so a notice whose
brand line items (AXIOM, XRY, GrayKey, ...) sit further down the description or only in a SOW attachment
never showed them in the report. 19AQMM27NITLABFURN (State AQM Momentum, INL Mexico lab/IT/furniture)
was the miss: the brands were in the full description ("MSAB, Magnet Axiom") and the IT SOW line items.

Read-only: GET requests to the same public sam.gov endpoints the website uses. No key, no writes to SAM.
"""
from __future__ import annotations

import html
import io
import json
import re
import time
import urllib.error
import urllib.request
import zipfile
from datetime import datetime, timezone
from typing import Any, Callable

OPP_URL = "https://sam.gov/api/prod/opps/v2/opportunities/{id}"
RESOURCES_URL = "https://sam.gov/api/prod/opps/v3/opportunities/{id}/resources"
FILE_URL = "https://sam.gov/api/prod/opps/v3/opportunities/resources/files/{rid}/download"
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
MAX_ATTACHMENT_BYTES = 8 * 1024 * 1024
MAX_ATTACHMENTS = 6
SCAN_VERSION = 1
RESCAN_DAYS = 5

# Canonical name -> regex body. Matched case-insensitively with word boundaries (no letter/digit either side).
# Longer names come first so "Magnet AXIOM" wins over bare "AXIOM" for the same span.
BRANDS: tuple[tuple[str, str], ...] = (
    ("Magnet AXIOM", r"magnet\s+axiom"),
    ("Magnet Forensics", r"magnet\s+forensics"),
    ("Magnet Verify", r"magnet\s+verify"),
    ("Magnet Review", r"magnet\s+review"),
    ("GrayKey", r"gr[ae]y\s?key"),          # "GreyKey" appears in the 19AQMM27NITLABFURN IT SOW
    ("Griffeye", r"griffeye"),
    ("MSAB", r"msab"),
    ("XRY", r"xry"),
    ("XAMN", r"xamn"),
    ("XEC", r"xec"),
    ("AXIOM", r"axiom"),
)
# Bare AXIOM and XEC are ordinary words/acronyms elsewhere (Axiom Builders, "axiom" in math, XEC as a part
# number). They count only when a forensic context word sits within CONTEXT_CHARS, or another brand matched.
NEEDS_CONTEXT = {"AXIOM", "XEC"}
CONTEXT_CHARS = 250
_CONTEXT = re.compile(
    r"(?<![a-z0-9])(forensics?|digital evidence|mobile|extraction|examiner|cellebrite|ftk|magnet|msab|xry|"
    r"graykey|greykey|license|licen[cs]es?|software|exploitation|e-?discovery|investigat\w*|cyber)(?![a-z0-9])",
    re.I,
)
_RES = tuple((name, re.compile(r"(?<![A-Za-z0-9])" + body + r"(?![A-Za-z0-9])", re.I)) for name, body in BRANDS)
_TAG = re.compile(r"<[^>]+>")


def html_to_text(body: str) -> str:
    return " ".join(html.unescape(_TAG.sub(" ", body or "")).split())


def find_brands(text: str) -> list[str]:
    """Canonical brand names found in text, in BRANDS order. Bare AXIOM/XEC need forensic context nearby."""
    text = text or ""
    taken: list[tuple[int, int]] = []
    found: dict[str, bool] = {}
    for name, rx in _RES:
        for m in rx.finditer(text):
            if any(s <= m.start() < e for s, e in taken):
                continue                  # inside a longer brand already matched ("Magnet AXIOM")
            taken.append((m.start(), m.end()))
            ok = True
            if name in NEEDS_CONTEXT:
                left = text[max(0, m.start() - CONTEXT_CHARS): m.start()]
                right = text[m.end(): m.end() + CONTEXT_CHARS]
                window = left + " " + right
                ok = bool(_CONTEXT.search(window))
            found[name] = found.get(name, False) or ok
    strong = [n for n, ok in found.items() if ok and n not in NEEDS_CONTEXT]
    out = []
    for name, _ in BRANDS:
        if name not in found:
            continue
        if found[name] or (name in NEEDS_CONTEXT and strong):
            out.append(name)
    if "Magnet AXIOM" in out and "AXIOM" in out:
        out.remove("AXIOM")                   # same product, named both ways
    return out


def _get(url: str, *, accept: str = "application/hal+json", timeout: float = 45.0, max_bytes: int | None = None) -> bytes:
    req = urllib.request.Request(url, headers={"Accept": accept, "User-Agent": UA, "Accept-Encoding": "identity"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        if max_bytes is None:
            return resp.read()
        data = resp.read(max_bytes + 1)
        if len(data) > max_bytes:
            raise ValueError("attachment too large")
        return data


def docx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = [n for n in z.namelist() if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)]
        xml = " ".join(z.read(n).decode("utf-8", "replace") for n in parts)
    return html_to_text(xml.replace("</w:p>", " \n"))


def xlsx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = [n for n in z.namelist() if n == "xl/sharedStrings.xml" or re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
        xml = " ".join(z.read(n).decode("utf-8", "replace") for n in names)
    return html_to_text(xml)


def pdf_text(data: bytes) -> str | None:
    try:
        from pypdf import PdfReader  # optional; not in requirements.txt today
    except ImportError:
        return None
    reader = PdfReader(io.BytesIO(data))
    return " ".join((p.extract_text() or "") for p in reader.pages[:60])


def attachment_text(name: str, data: bytes) -> str | None:
    low = (name or "").lower()
    if low.endswith(".docx"):
        return docx_text(data)
    if low.endswith((".xlsx", ".xlsm")):
        return xlsx_text(data)
    if low.endswith(".pdf"):
        return pdf_text(data)
    if low.endswith((".txt", ".csv")):
        return data.decode("utf-8", "replace")
    return None                               # .doc/.xls/.zip etc. are not read


def fetch_description(notice_id: str, get: Callable[..., bytes] = _get) -> tuple[str, str]:
    """-> (plain description text, modifiedDate)."""
    doc = json.loads(get(OPP_URL.format(id=notice_id)).decode("utf-8"))
    descs = doc.get("description") or []
    body = " ".join(str(d.get("body") or "") for d in descs if isinstance(d, dict))
    return html_to_text(body), str(doc.get("modifiedDate") or "")


def fetch_attachments(notice_id: str, get: Callable[..., bytes] = _get, sleep: float = 1.0) -> tuple[list[tuple[str, str]], list[str]]:
    """-> ([(file name, text)], [skipped file names]). Public attachments only."""
    doc = json.loads(get(RESOURCES_URL.format(id=notice_id)).decode("utf-8"))
    files = []
    for grp in ((doc.get("_embedded") or {}).get("opportunityAttachmentList") or []):
        for a in grp.get("attachments") or []:
            if a.get("type") == "file" and a.get("accessLevel") == "public" and str(a.get("deletedFlag") or "0") == "0":
                files.append(a)
    out, skipped = [], []
    for a in files[:MAX_ATTACHMENTS]:
        name = str(a.get("name") or "")
        if int(a.get("size") or 0) > MAX_ATTACHMENT_BYTES:
            skipped.append(name)
            continue
        try:
            data = get(FILE_URL.format(rid=a.get("resourceId")), accept="*/*", max_bytes=MAX_ATTACHMENT_BYTES)
            text = attachment_text(name, data)
        except Exception:  # noqa: BLE001 - one bad file must not stop the run
            text = None
        if text is None:
            skipped.append(name)
        else:
            out.append((name, text))
        if sleep:
            time.sleep(sleep)
    skipped += [str(a.get("name") or "") for a in files[MAX_ATTACHMENTS:]]
    return out, skipped


def scan_notice(rec: dict[str, Any], *, attachments: bool = False, get: Callable[..., bytes] = _get,
                sleep: float = 1.0) -> dict[str, Any]:
    """Brand fields for one history record. Title + full description always; attachments when asked."""
    nid = str(rec.get("notice_id") or "")
    terms: dict[str, list[str]] = {}

    def add(found: list[str], where: str) -> None:
        for name in found:
            terms.setdefault(name, [])
            if where not in terms[name]:
                terms[name].append(where)

    add(find_brands(str(rec.get("title") or "")), "title")
    desc, modified = fetch_description(nid, get)
    add(find_brands(desc), "description")
    skipped: list[str] = []
    if attachments:
        files, skipped = fetch_attachments(nid, get, sleep=sleep)
        for _name, text in files:
            add(find_brands(text), "attachment")
    order = [n for n, _ in BRANDS]
    names = sorted(terms, key=order.index)
    return {
        "brand_terms": names,
        "brand_where": {n: terms[n] for n in names},
        "brand_scan": {"version": SCAN_VERSION, "modified": modified, "attachments": bool(attachments),
                       "skipped": skipped[:10], "at": datetime.now(timezone.utc).isoformat(timespec="seconds")},
    }


def needs_scan(rec: dict[str, Any], *, attachments: bool) -> bool:
    prev = rec.get("brand_scan")
    if not isinstance(prev, dict) or prev.get("version") != SCAN_VERSION:
        return True
    if attachments and not prev.get("attachments"):
        return True
    if rec.get("brand_terms"):
        return False                          # a brand match sticks
    # No match yet: look again after RESCAN_DAYS, so an amendment that adds a SOW is picked up.
    try:
        at = datetime.fromisoformat(str(prev.get("at") or ""))
    except ValueError:
        return True
    return (datetime.now(timezone.utc) - at).days >= RESCAN_DAYS


def scan_history(history: dict[str, Any], keys: list[str], *, attachments: bool = False, sleep: float = 1.0,
                 limit: int = 400, get: Callable[..., bytes] = _get, log: Callable[[str], None] = print) -> dict[str, int]:
    """Scan the given notice keys once (cached in the record). Errors keep the record unscanned for next run."""
    stats = {"scanned": 0, "cached": 0, "errors": 0, "brand": 0}
    notices = history.get("notices") or {}
    for key in keys:
        rec = notices.get(key)
        if not isinstance(rec, dict) or not rec.get("notice_id"):
            continue
        if not needs_scan(rec, attachments=attachments):
            stats["cached"] += 1
            stats["brand"] += bool(rec.get("brand_terms"))
            continue
        if stats["scanned"] >= limit:
            break
        try:
            rec.update(scan_notice(rec, attachments=attachments, get=get, sleep=sleep))
            stats["scanned"] += 1
            stats["brand"] += bool(rec.get("brand_terms"))
            if rec.get("brand_terms"):
                log(f"  brand match {rec.get('solicitation_number') or key}: {', '.join(rec['brand_terms'])}")
        except (urllib.error.URLError, ValueError, OSError, json.JSONDecodeError) as e:
            stats["errors"] += 1
            log(f"  brand scan failed for {key}: {type(e).__name__}")
        if sleep:
            time.sleep(sleep)
    return stats
