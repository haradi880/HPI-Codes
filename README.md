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

Each scraper writes JSON, XLSX, raw exports, API call detail, and API report files inside its category folder.
