from __future__ import annotations

from collections import OrderedDict

import pandas as pd
import pytest

import sec_nongaap as ng


SEC_SOURCE = {
    "role": "Press release",
    "document": "earnings-release.htm",
    "description": "Earnings release",
    "url": "https://www.sec.gov/Archives/edgar/data/1/1/earnings-release.htm",
}


def test_extract_metric_definitions_keeps_evidence_and_source() -> None:
    text = (
        "Adjusted EBITDA is defined as net income before interest, income taxes, "
        "depreciation and amortization, excluding restructuring charges."
    )

    definitions = ng.extract_metric_definitions(text, SEC_SOURCE)

    assert len(definitions) == 1
    definition = definitions[0]
    assert definition["metric"] == "Adjusted EBITDA"
    assert definition["definition_type"] == "Calculation or composition"
    assert "restructuring charges" in definition["definition_context"]
    assert definition["source_url"] == SEC_SOURCE["url"]


def test_extract_metric_definitions_does_not_promote_a_plain_metric_mention() -> None:
    definitions = ng.extract_metric_definitions(
        "Adjusted EBITDA was $42 million in the quarter.",
        SEC_SOURCE,
    )

    assert definitions == []


def test_relevant_exhibits_accepts_only_sec_hosted_documents() -> None:
    documents = [
        {
            "doc_type": "EX-99.1",
            "role": "Press release",
            "url": SEC_SOURCE["url"],
            "sequence": "1",
        },
        {
            "doc_type": "EX-99.2",
            "role": "Press release",
            "url": "https://untrusted.example/press-release.htm",
            "sequence": "2",
        },
    ]

    exhibits = ng.relevant_exhibits(documents)

    assert [item["url"] for item in exhibits] == [SEC_SOURCE["url"]]


def test_primary_8k_link_discovery_finds_only_sec_press_release_links() -> None:
    primary_html = """
    <html><body>
      <p>Our <a href="earnings-release.pdf">quarterly earnings press release</a> is furnished as Exhibit 99.1.</p>
      <a href="https://untrusted.example/earnings-release.pdf">external release</a>
      <a href="governance.htm">corporate governance</a>
    </body></html>
    """

    documents = ng.discover_primary_document_links(
        primary_html,
        "https://www.sec.gov/Archives/edgar/data/1/1/primary8k.htm",
    )

    assert len(documents) == 1
    assert documents[0]["url"].endswith("/earnings-release.pdf")
    assert documents[0]["role"] == "Press release"
    assert documents[0]["discovery_method"] == "8-K primary link"


def test_sec_client_rejects_non_sec_resource_before_request() -> None:
    client = ng.SecClient("research@example.com")

    with pytest.raises(ValueError, match="SEC"):
        client.get_bytes("https://untrusted.example/document.htm")


def test_sec_client_evicts_old_documents_from_bounded_cache() -> None:
    client = ng.SecClient("research@example.com", max_cache_bytes=5)
    one = ng.FetchedResource(b"123", "text/html", "https://www.sec.gov/one.htm")
    two = ng.FetchedResource(b"456", "text/html", "https://www.sec.gov/two.htm")

    client._cache_resource(one.url, one)
    client._cache_resource(two.url, two)

    assert isinstance(client._cache, OrderedDict)
    assert list(client._cache) == [two.url]
    assert client._cache_bytes == len(two.content)


def test_sec_client_records_and_raises_after_rate_limit_retries(monkeypatch: pytest.MonkeyPatch) -> None:
    class RateLimitedResponse:
        status_code = 429
        headers = {"Retry-After": "0"}
        url = SEC_SOURCE["url"]

    client = ng.SecClient("research@example.com")
    monkeypatch.setattr(client.session, "get", lambda *_args, **_kwargs: RateLimitedResponse())
    monkeypatch.setattr(ng.time, "sleep", lambda *_args: None)

    with pytest.raises(ng.SecRateLimitError):
        client.get_bytes(SEC_SOURCE["url"])

    events = client.request_events_frame()
    assert len(events) == 4
    assert set(events["category"]) == {"SEC EDGAR rate limit"}
    assert events.iloc[-1]["severity"] == "Error"
    assert events.iloc[-1]["status_code"] == 429


def test_missing_earnings_release_is_a_warning_and_request_event(monkeypatch: pytest.MonkeyPatch) -> None:
    client = ng.SecClient("research@example.com")
    anchors = pd.DataFrame(
        [
            {
                "fiscal_year": 2026,
                "fiscal_quarter": "Q2",
                "period_end": "2026-06-30",
                "periodic_form": "10-Q",
                "periodic_filing_date": "2026-08-01",
                "periodic_document": "quarterly.htm",
                "periodic_url": SEC_SOURCE["url"],
            }
        ]
    )
    monkeypatch.setattr(ng, "match_earnings_8k", lambda *_args, **_kwargs: None)

    analysis = ng.analyze_company_quarters(client, 1, pd.DataFrame(), anchors, [2026])

    assert analysis["coverage"].iloc[0]["status"] == "No matching earnings 8-K found"
    assert analysis["warnings"].iloc[0]["warning_category"] == "Missing earnings release"
    assert analysis["request_events"].iloc[0]["category"] == "Missing earnings release"


def test_export_metrics_table_retains_reconciliation_and_source_fields() -> None:
    reconciliations = pd.DataFrame(
        [
            {
                "pair_id": "pair-1",
                "period": "FY2026 Q2",
                "period_end": "2026-06-30",
                "metric": "Adjusted EBITDA",
                "gaap_label": "Net income",
                "gaap_display": "$10.0",
                "adjustment_display": "$4.0",
                "non_gaap_label": "Adjusted EBITDA",
                "non_gaap_display": "$14.0",
                "unit": "usd",
                "scale": "millions",
                "confidence": "High",
                "source_role": "Press release",
                "source_page": 3,
                "source_url": SEC_SOURCE["url"],
            }
        ]
    )
    tieouts = pd.DataFrame(
        [{"pair_id": "pair-1", "tie_out_status": "Ties within rounding", "tie_out_note": "Difference: $0.0"}]
    )

    export = ng.make_export_metrics_table(reconciliations, tieouts)

    assert export.to_dict("records") == [
        {
            "Fiscal period": "FY2026 Q2",
            "Period end": "2026-06-30",
            "Non-GAAP metric": "Adjusted EBITDA",
            "Comparable GAAP label": "Net income",
            "Comparable GAAP value": "$10.0",
            "Total adjustments": "$4.0",
            "Reported non-GAAP label": "Adjusted EBITDA",
            "Reported non-GAAP value": "$14.0",
            "Unit": "usd",
            "Scale": "millions",
            "Parse confidence": "High",
            "Tie-out status": "Ties within rounding",
            "Tie-out note": "Difference: $0.0",
            "Source type": "Press release",
            "PDF page": 3,
            "SEC source": SEC_SOURCE["url"],
        }
    ]


def test_peer_adjustment_comparison_matrix_keeps_latest_exact_labels_side_by_side() -> None:
    adjustments = pd.DataFrame(
        [
            {
                "company": "AAPL",
                "fiscal_year": 2026,
                "fiscal_quarter": "Q2",
                "period": "FY2026 Q2",
                "adjustment_category": "Stock-based and equity compensation",
                "adjustment_label": "Share-based compensation",
                "adjustment_display": "$8",
            },
            {
                "company": "AAPL",
                "fiscal_year": 2026,
                "fiscal_quarter": "Q3",
                "period": "FY2026 Q3",
                "adjustment_category": "Stock-based and equity compensation",
                "adjustment_label": "Share-based compensation",
                "adjustment_display": "$12",
            },
            {
                "company": "MSFT",
                "fiscal_year": 2026,
                "fiscal_quarter": "Q4",
                "period": "FY2026 Q4",
                "adjustment_category": "Stock-based and equity compensation",
                "adjustment_label": "Stock-based compensation expense",
                "adjustment_display": "$7",
            },
        ]
    )

    matrix = ng.make_peer_adjustment_comparison_matrix(
        adjustments,
        latest_period_only=True,
        minimum_companies=2,
    )

    assert list(matrix.columns) == [
        "Adjustment category",
        "Peers disclosing",
        "AAPL\nFY2026 Q3",
        "MSFT\nFY2026 Q4",
    ]
    assert matrix.iloc[0]["Peers disclosing"] == 2
    assert matrix.iloc[0]["AAPL\nFY2026 Q3"] == "Share-based compensation ($12)"
    assert matrix.iloc[0]["MSFT\nFY2026 Q4"] == "Stock-based compensation expense ($7)"
