# HPI Codes

Upload-ready HPI scraper package.

## Folder

`hpi_scraper/` contains the five category scrapers:

- `build_firmographic_structure.py`
- `build_jobs_hiring_structure.py`
- `build_technographic_structure.py`
- `news_announcements_fetch.py`
- `build_contact_level_structure.py`

The package also includes the latest generated Sea Limited data/records under the category folders.

## Input

Edit `hpi_scraper/input/compnys.txt`.

The file can contain 10, 50, or more companies as long as it keeps these columns:

```csv
source_rank,company_name,domain,linkedin_url,source_basis
```

## Secrets

Do not commit real keys. Copy `.env.example` to `.env` locally and fill in keys before running.

## Run

```bash
cd hpi_scraper
pip install -r requirements.txt
python run_all.py
```

The uploaded sample keeps only Excel outputs (`.xlsx`) plus runnable code. JSON/raw API exports are intentionally excluded.

## Sea Limited Output Notes

Some Excel fields are blank because the provider response did not include those fields. For example, Apollo and Coresignal do not return every firmographic field, and TheirStack/Coresignal technographic responses do not return IT spend.

Contact-level has no matched contacts in the current sample because FullEnrich returned `not_enough_credits`, and the Apollo people-search fallback is not available on the current Apollo plan. The API report workbook records that blocker.
