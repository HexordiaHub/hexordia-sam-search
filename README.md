# SAM daily search

Automated SAM.gov opportunity keyword checks. Produces Excel + HTML reports.

## Public report

After GitHub Pages is enabled and a workflow has run:

- HTML: https://hexordiahub.github.io/hexordia-sam-search/v/
- Excel: https://hexordiahub.github.io/hexordia-sam-search/v/SAM-daily-latest.xlsx

Low-profile path (`/v/`), `noindex`, robots disallow. Content is public SAM.gov notice metadata only.

## GitHub Actions (once)

1. *(Optional)* Secret `SAM_API_KEY` — not required for Actions (frontend only)  
2. **Settings → Pages → Source:** GitHub Actions  
3. **Settings → Actions → General → Workflow permissions:** Read and write  
4. **Actions → SAM daily search → Run workflow** (first run)

Schedule: `0 9 * * *` UTC (≈ 04:00 EST / 05:00 EDT). **Currently disabled** in `.github/workflows/sam-daily.yml` after the migration from the old public repo; re-enable after a successful manual run.

CI uses **moderate pacing** (`--robot`, ~3s between terms — typically a few minutes, not ~20). Local runs can use slower human pacing if you prefer.

## Local (optional)

```bash
git clone https://github.com/HexordiaHub/hexordia-sam-search.git
cd hexordia-sam-search
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
cp .env.example .env   # add SAM_API_KEY if using API mode
.venv/bin/python scripts/sam_search.py --days 1 --source both --no-latest-sync
```

Terms: `config/search_terms.json`.

## History retention

Reports keep a **rolling 15 days** of day buckets and notices (`--history-days 15`).
Each scheduled search uses a **15-day posted-date window** (`--days 15`) so the tracker accumulates a two-week record.


## Long-running archive vs 15-day webpage

| What | Where | Retention |
|------|--------|-----------|
| Public webpage / Excel on Pages | `docs/v/` | **Rolling 15 days** |
| Durable notice list | `data/archive/notices-master.json` + `.csv` | **Keeps growing** (not purged) |
| Per-day snapshots | `data/archive/days/YYYY-MM-DD.json` | Kept in git |
| Append audit log | `data/archive/notices-append.jsonl` | Append-only |

The site stays lean; the repo keeps the long history under `data/archive/` (not deployed to Pages).

## Copy for Trello

Each result row on the HTML report has **Copy for Trello**. Click to copy a card-ready block (title, SAM link, terms, deadline, etc.), then paste into a new Trello card. No Trello API keys required.


## Brand-name scan

After each search, `scripts/brand_scan.py` reads every notice's **full description** (one GET to the public
`sam.gov/api/prod/opps/v2/opportunities/<id>` per notice, cached in `data/history.json` under `brand_scan`) and records
`brand_terms` / `brand_where` for AXIOM, Magnet AXIOM, Magnet Forensics, Magnet Verify, Magnet Review, GrayKey (also
"GreyKey"), Griffeye, MSAB, XRY, XAMN, XEC. Case-insensitive, word boundaries; bare AXIOM and XEC count only with
forensic context nearby (or another brand in the notice), so "Axiom Builders" or "axiom" in math is not a hit.
Unmatched notices are rechecked after 5 days. `--no-brand-scan` turns it off.

`--brand-attachments` (or `SAM_BRAND_ATTACHMENTS=1`) also reads up to 6 public attachments per notice (≤ 8 MB):
.docx/.xlsx/.txt with the standard library, .pdf only when `pypdf` is installed. Off by default; the workflow does
not pass it today. The Excel report has a **Brand match** column; the Hexordia Hub reads `brand_terms` from history.json.

Tests: `python3 tests/test_brand_scan.py`
