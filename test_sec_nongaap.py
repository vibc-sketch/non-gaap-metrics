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
