# SEC Non-GAAP Metrics Explorer

A Streamlit app that searches public-company issuers and extracts non-GAAP metric disclosures from SEC filings and exhibits.

## What it does

- Search issuer universe by company name or ticker; optionally constrain by SIC code.
- Resolve issuer metadata from SEC submissions, including SIC and fiscal year end.
- Inspect 8-K / 8-K-A, 10-K / 10-Q / amendments, and optionally S-1 / S-1-A filings.
- Find likely earnings press-release exhibits (especially EX-99.x) and parse tables/narrative for common non-GAAP disclosures.
- Also inspect SEC-hosted PDF/HTML press-release links embedded in the primary earnings 8-K when an exhibit-table entry is incomplete or absent.
- Normalize records to issuer fiscal year / fiscal quarter using SEC filing `fy` / `fp` metadata rather than assuming calendar quarters.
- Preserve issuer-provided definition and calculation evidence when an exhibit explicitly describes a non-GAAP measure.
- Build source-linked reconciliation bridges, adjustment histories, peer disclosure matrices, and Excel/CSV exports.
- Compare selected peer companies in a custom side-by-side adjustment matrix that aligns normalized categories while retaining the issuer's exact labels, values, fiscal periods, and SEC source rows.
- Chart selected peer adjustment categories as fiscal-period disclosure-presence trends, avoiding aggregation of non-comparable issuer-reported adjustment amounts.
- Accept peer tickers, company names, recognizable name fragments, and common legal-name suffix variants when resolving the peer set against SEC company data.
- Export a presentation-style reconciliation workbook with issuer-reported GAAP-to-non-GAAP bridge hierarchy, SEC source links, blue source-input values, and cell-level audit comments.
- Use an Apple-inspired interface with forced light data surfaces, explicit sidebar/header/tab/alert/table colors, and high-contrast controls that do not inherit browser or operating-system dark-mode foregrounds.
- Pause and retry automatically when EDGAR returns a rate limit or temporary service error, while recording the recovery or failure in the Source audit log.
- Flag periods where no matching earnings 8-K or SEC-hosted press-release link was identified, rather than silently returning an empty result.
- Export extracted records to CSV with direct SEC source links.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Test

```bash
pip install -r requirements-dev.txt
pytest -q
```

## Production notes

1. Replace the default SEC User-Agent with a real contact identity in `app.py` (or externalize it into configuration).
2. Add a persistent issuer/SIC index for faster industry searches.
3. Extend the test fixture library across issuer-specific table formats, text-heavy PDFs, image-only PDFs, and non-calendar fiscal years.
4. Add a database (Postgres/DuckDB) and a background ingestion job for a multi-user deployment.
