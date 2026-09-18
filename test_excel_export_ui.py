from __future__ import annotations

import pandas as pd
from streamlit.testing.v1 import AppTest


def test_completed_analysis_exposes_metrics_and_reconciliations_excel_download() -> None:
    app = AppTest.from_file("app.py", default_timeout=20)
    app.run()

    app.session_state.company = {
        "name": "Example Issuer, Inc.",
        "ticker": "EXMP",
        "cik": 1,
        "exchange": "NASDAQ",
        "sic": "7372",
        "sic_description": "Prepackaged Software",
        "fiscal_year_end": "1231",
    }
    app.session_state.anchors = pd.DataFrame(
        [
            {
                "fiscal_year": 2026,
                "fiscal_quarter": "Q2",
                "period_end": "2026-06-30",
                "periodic_form": "10-Q",
                "periodic_filing_date": "2026-08-01",
                "metadata_source": "Inline XBRL",
                "periodic_url": "https://www.sec.gov/Archives/edgar/data/1/1/quarterly.htm",
            }
        ]
    )
    app.session_state.analysis_years = [2026]
    app.session_state.analysis = {
        "coverage": pd.DataFrame(),
        "reconciliations": pd.DataFrame(
            [
                {
                    "pair_id": "pair-1",
                    "period": "FY2026 Q2",
                    "fiscal_year": 2026,
                    "fiscal_quarter": "Q2",
                    "period_end": "2026-06-30",
                    "metric": "Adjusted EBITDA",
                    "gaap_label": "Net income",
                    "gaap_display": "$10.0",
                    "gaap_value": 10.0,
                    "adjustment_display": "$4.0",
                    "adjustment_value": 4.0,
                    "non_gaap_label": "Adjusted EBITDA",
                    "non_gaap_display": "$14.0",
                    "non_gaap_value": 14.0,
                    "unit": "usd",
                    "scale": "millions",
                    "confidence": "High",
                    "source_role": "Press release",
                    "source_page": 1,
                    "source_url": "https://www.sec.gov/Archives/edgar/data/1/1/earnings.htm",
                    "table_title": "Non-GAAP reconciliation",
                }
            ]
        ),
        "adjustments": pd.DataFrame(),
        "adjustment_history": pd.DataFrame(),
        "adjustment_tieouts": pd.DataFrame(),
        "mentions": pd.DataFrame(),
        "definitions": pd.DataFrame(),
        "kpis": pd.DataFrame(),
        "sources": pd.DataFrame(),
        "evidence": pd.DataFrame(),
        "warnings": pd.DataFrame(),
        "request_events": pd.DataFrame(),
    }

    app.run()

    assert not app.exception
    download_buttons = app.get("download_button")
    labels = [element.label for element in download_buttons]
    assert "Download metrics & reconciliations Excel" in labels
    assert "Download CSV package" in labels
