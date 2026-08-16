# SEC Non-GAAP Metrics Explorer

A Streamlit app that searches public-company issuers and extracts non-GAAP metric disclosures from SEC filings and exhibits.

## What it does

- Search issuer universe by company name or ticker; optionally constrain by SIC code.
- Resolve issuer metadata from SEC submissions, including SIC and fiscal year end.
- Inspect 8-K / 8-K-A, 10-K / 10-Q / amendments, and optionally S-1 / S-1-A filings.
- Find likely earnings press-release exhibits (especially EX-99.x) and parse tables/narrative for common non-GAAP disclosures.
- Normalize records to issuer fiscal year / fiscal quarter using SEC filing `fy` / `fp` metadata rather than assuming calendar quarters.
- Export extracted records to CSV with direct SEC source links.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Production notes

1. Replace the default SEC User-Agent with a real contact identity in `app.py` (or externalize it into configuration).
2. Add a persistent issuer/SIC index for faster industry searches.
3. Add stronger table parsing and metric dictionaries per issuer to distinguish metric labels, values, and units.
4. Add an audit layer that stores exact source exhibit, page/table, and extraction confidence.
5. Add a database (Postgres/DuckDB) and a background ingestion job for a multi-user deployment.
