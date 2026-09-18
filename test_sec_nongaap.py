from __future__ import annotations

from collections import OrderedDict

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
