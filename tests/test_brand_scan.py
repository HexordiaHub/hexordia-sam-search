"""Brand-name scan (scripts/brand_scan.py) and the multi-word body-term fix in sam_search.py. No network.
Usage: python3 tests/test_brand_scan.py"""
import io
import json
import os
import sys
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
import brand_scan  # noqa: E402
import sam_search  # noqa: E402

LABFURN_ID = "f7d684954d7147f0a993161e6f58d9d1"
with open(os.path.join(HERE, "fixtures", "19AQMM27NITLABFURN-description.txt"), encoding="utf-8") as fh:
    LABFURN_DESC = fh.read()
# Line items as they appear in the IT SOW attachment (SOW - IT Equipment - Draft.docx).
IT_SOW = ("6AX210 Magnet Axiom Magnet Axiom Advanced 1 year of software license Brand Name Only 1 Ea. "
          "GreyKey Software Suite 1 Ea. MSAB XRY Pro Suite 1 Ea. Mobile Forensics and Data Recovery Software for "
          "Two Years MSAB XRY EXTRACT and XAMN Brand Name Only")


def docx(text):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", "<w:document><w:body><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:body></w:document>" % text)
    return buf.getvalue()


def fake_get(desc_html, files=()):
    def get(url, **kw):
        if url.endswith("/resources"):
            atts = [{"type": "file", "accessLevel": "public", "deletedFlag": "0", "name": n, "resourceId": "r%d" % i,
                     "size": len(b)} for i, (n, b) in enumerate(files)]
            return json.dumps({"_embedded": {"opportunityAttachmentList": [{"attachments": atts}]}}).encode()
        if "/files/" in url:
            rid = int(url.split("/files/r")[1].split("/")[0])
            return files[rid][1]
        return json.dumps({"modifiedDate": "2026-10-05T18:00:46Z", "description": [{"body": desc_html}]}).encode()
    return get


class FindBrands(unittest.TestCase):
    def test_list_and_boundaries(self):
        cases = {
            "e.g., FTK, Cellebrite, MSAB, Magnet Axiom": ["Magnet AXIOM", "MSAB"],
            IT_SOW: ["Magnet AXIOM", "GrayKey", "MSAB", "XRY", "XAMN"],
            "Magnet Forensics GrayKey and Magnet Verify, Magnet Review": ["Magnet Forensics", "Magnet Verify", "Magnet Review", "GrayKey"],
            "Griffeye Analyze DI Pro": ["Griffeye"],
            "graykey GRAYKEY Gray Key": ["GrayKey"],
            "MSABX, XRYLENE, AXIOMATIC, magnetic review": [],
        }
        for text, want in cases.items():
            self.assertEqual(brand_scan.find_brands(text), want, text)

    def test_generic_axiom_and_xec_need_forensic_context(self):
        self.assertEqual(brand_scan.find_brands("Award to AXIOM BUILDERS INC for HVAC refurbishment"), [])
        self.assertEqual(brand_scan.find_brands("By the axiom of choice every vector space has a basis."), [])
        self.assertEqual(brand_scan.find_brands("Replace XEC-200 valve assembly on chiller 3"), [])
        self.assertEqual(brand_scan.find_brands("AXIOM Cyber license for the digital forensics lab"), ["AXIOM"])
        self.assertEqual(brand_scan.find_brands("XEC Director for mobile device extraction"), ["XEC"])
        # Another brand in the same notice is context enough.
        self.assertEqual(brand_scan.find_brands("Axiom renewal. Separately: XRY."), ["XRY", "AXIOM"])

    def test_labfurn_description_alone_is_a_brand_match(self):
        self.assertEqual(brand_scan.find_brands(LABFURN_DESC), ["Magnet AXIOM", "MSAB"])


class ScanNotice(unittest.TestCase):
    def test_description_then_attachments(self):
        rec = {"notice_id": LABFURN_ID, "title": "IT, Lab, and Furniture Equipment"}
        get = fake_get("<p>%s</p>" % LABFURN_DESC, files=[("SOW - IT Equipment - Draft.docx", docx(IT_SOW)),
                                                         ("Map.dwg", b"x")])
        desc_only = brand_scan.scan_notice(rec, attachments=False, get=get, sleep=0)
        self.assertEqual(desc_only["brand_terms"], ["Magnet AXIOM", "MSAB"])
        full = brand_scan.scan_notice(rec, attachments=True, get=get, sleep=0)
        self.assertEqual(full["brand_terms"], ["Magnet AXIOM", "GrayKey", "MSAB", "XRY", "XAMN"])
        self.assertEqual(full["brand_where"]["XRY"], ["attachment"])
        self.assertEqual(full["brand_where"]["MSAB"], ["description", "attachment"])
        self.assertEqual(full["brand_scan"]["skipped"], ["Map.dwg"])

    def test_scan_history_caches_and_survives_errors(self):
        hist = {"notices": {"a": {"notice_id": "a", "title": "x"}, "b": {"notice_id": "b", "title": "y"}}}
        calls = []

        def get(url, **kw):
            calls.append(url)
            if "/b" in url:
                raise OSError("down")
            return json.dumps({"description": [{"body": "Magnet Forensics GrayKey"}]}).encode()
        st = brand_scan.scan_history(hist, ["a", "b"], get=get, sleep=0, log=lambda _m: None)
        self.assertEqual(st, {"scanned": 1, "cached": 0, "errors": 1, "brand": 1})
        self.assertNotIn("brand_scan", hist["notices"]["b"])
        st = brand_scan.scan_history(hist, ["a"], get=get, sleep=0, log=lambda _m: None)
        self.assertEqual(st["cached"], 1)
        self.assertEqual(len([c for c in calls if c.endswith("/a")]), 1)


class BodyTerms(unittest.TestCase):
    def row(self, title, desc):
        return {"_id": "n1", "title": title, "publishDate": "2026-10-10", "descriptions": [{"content": desc}],
                "type": {"value": "Solicitation"}, "isActive": True}

    def test_multi_word_phrase_in_description_survives_merge_and_report(self):
        h = sam_search.opportunity_from_frontend_row(self.row("Lab software", "Renew Magnet Forensics licenses"),
                                                     "Magnet Forensics")
        self.assertEqual(h.body_terms, {"Magnet Forensics"})
        merged = sam_search.merge_hits([h])
        self.assertEqual(len(merged), 1)
        rec = sam_search.hit_to_notice_dict(merged[0])
        self.assertEqual(rec["body_terms"], ["Magnet Forensics"])
        self.assertEqual(sam_search.filter_notice_row(rec)["matched_terms"], ["Magnet Forensics"])

    def test_phrase_nowhere_is_still_dropped(self):
        self.assertIsNone(sam_search.opportunity_from_frontend_row(self.row("Lab", "Magnet and forensics"), "Magnet Forensics"))
        self.assertIsNone(sam_search.filter_notice_row({"title": "Lab", "matched_terms": ["Digital forensics"]}))

    def test_brand_row_kept_in_report_without_title_terms(self):
        row = {"title": "Lab", "matched_terms": ["Digital forensics"], "brand_terms": ["XRY"]}
        self.assertEqual(sam_search.filter_notice_row(row)["matched_terms"], [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
