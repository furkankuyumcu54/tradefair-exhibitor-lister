# Tradefair Exhibitor Lister

Scrapes exhibitor/participant listings from Turkish trade fair websites and enriches the data with contact information (email, phone, website).

## How it works

1. **Scrape** — Uses Playwright (headless Chromium) to visit a fair's exhibitor listing page, paginate through all pages, and extract company details (name, website, phone, hall, stand).
2. **Enrich** — For each company missing contact info, searches DuckDuckGo and guesses domains to find the company website, then scrapes it for email addresses and phone numbers.

## Project structure

```
├── agent.py                # Main entry point — orchestrates scrape + enrich
├── enricher.py             # Contact info enrichment (DuckDuckGo, domain guessing, scraping)
├── requirements.txt        # Python dependencies
├── fairs/                  # YAML config files (one per trade fair)
│   ├── intermob.yaml
│   ├── woodtechistanbul.yaml
│   └── ...
├── scrapers/
│   └── tuyap.py            # Tuyap-specific Playwright scraper
└── output/                 # Generated Excel files
```

## Prerequisites

- Python 3.12+
- Chromium (installed via Playwright)

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
python agent.py fairs/<fair-name>.yaml
```

Example:

```bash
python agent.py fairs/intermob.yaml
```

This produces two files in `output/`:
- `<fair-slug>.xlsx` — raw scraped data
- `<fair-slug>-enriched.xlsx` — data with enriched contact info

The enricher can also be run standalone:

```bash
python enricher.py <input.xlsx> <output.xlsx>
```

## Configuration

Each fair has a YAML file in `fairs/`:

```yaml
fair:
  name: "Intermob 2026"
source:
  platform: tuyap
  exhibitors_url: "https://intermobistanbul.com/katilimci-listesi"
```

Currently only the `tuyap` platform is supported. All fairs use Tuyap's exhibitor listing system (`/katilimci-listesi`).

## Output format

| Column | Description |
|---|---|
| Company Name | Exhibitor name |
| Website | Company website |
| Phone | Phone number |
| Email | Email address |
| Hall | Exhibition hall |
| Stand | Stand number |
| Detail Page URL | Link to the exhibitor's detail page |
