# SEC Non-GAAP Metrics Explorer: Product and Engineering Audit

**Date:** September 17, 2026  
**Author:** Manus AI  
**Scope:** Review of the public Streamlit application, the `vibc-sketch/non-gaap-metrics` repository, and a local validation of targeted code improvements.

## Executive conclusion

The application has a credible, evidence-first foundation. It already performs the difficult parts of the intended workflow: issuer resolution, fiscal-period anchoring, earnings 8-K matching, exhibit extraction, GAAP-to-non-GAAP bridges, adjustment-category normalization, peer disclosure matrices, source links, and exports. Its strongest design choice is that it uses periodic reports for fiscal-period normalization while restricting metric extraction to the matched 8-K exhibit package.

The most important product gap was that users could see a measure and its reconciliation but not the issuer language explaining how the measure was defined or calculated. The most important engineering risk was that the application repeated expensive SEC work across runs while retaining every downloaded document indefinitely inside a mutable client cache. The implemented changes add a source-evidenced definition workflow, cache public analysis results for one hour, bound in-memory document caching, keep mutable HTTP sessions session-scoped, add SEC-host validation, and introduce automated tests.

A live interactive workflow could not be completed during this audit because the public site rendered a blank application canvas in the unauthenticated audit browser after the Streamlit platform shell loaded. The navigation record completed in 11.64 seconds, but the page body contained no visible app text. The platform status endpoint was reachable, so this is a user-facing availability or rendering observation rather than a confirmed application-code exception. The local Streamlit server started normally and returned a healthy status response.

## Implemented improvements

| Area | Change | User and operational effect |
| --- | --- | --- |
| **Definitions and methodology** | Added `extract_metric_definitions`, which retains only nearby issuer language containing a definition or calculation cue. Added a **Definitions & method** tab and included definitions in Excel and CSV exports. | Users can review the exact issuer text and SEC source behind a measure rather than treating a nearby metric mention as a standardized definition. |
| **Repeated-work performance** | Added one-hour `st.cache_data` caches for issuer metadata, fiscal anchors, completed issuer analyses, and completed peer analyses. | Reruns with the same issuer, fiscal years, and exhibit depth avoid refetching and reparsing public filing data. Peer analyses can be reused instead of repeating the full pipeline. |
| **Memory control** | Replaced the unlimited per-client document dictionary with a least-recently-used cache capped at 160 MB. | Multi-exhibit and multi-peer runs no longer allow document bytes to grow without bound within a browser session. |
| **Session safety** | Changed the mutable `requests.Session`/SEC client from globally shared `st.cache_resource` storage to browser-session state. | One user’s HTTP session and in-process document cache are not shared with other users. This aligns with Streamlit’s requirement that globally cached resources be thread-safe. [3] |
| **Source hardening** | Allowed retrieval and rendered source buttons only for HTTPS `sec.gov`, `www.sec.gov`, and `data.sec.gov` URLs. Redirects outside those hosts now fail closed. | Source links and document requests cannot be redirected to arbitrary hosts through parsed filing metadata. |
| **Regression protection** | Added focused engine tests and a Streamlit initial-render smoke test. | Definition evidence, source-host validation, cache eviction, and the initial application state now have repeatable automated coverage. |

## How the improved workflow maps to the intended product

The revised workflow supports the product intent in a clear order. First, the user selects an issuer and a fiscal-period window. The app identifies the relevant earnings 8-K and exhibit package. It then produces a catalog of detected non-GAAP measures, a period-by-period bridge from the comparable GAAP result through issuer-reported adjustment items to the non-GAAP result, and a definition-evidence view. The definition view is intentionally conservative: it stores exact text only when a nearby phrase appears to define, calculate, describe composition, or state exclusions for the measure. The interface labels this as issuer evidence rather than claiming a universally comparable definition.

Peer benchmarking continues to normalize each company to its own fiscal calendar. It compares disclosure presence and adjustment categories, while detailed peer bridges preserve the issuer’s actual labels. This is a sound approach because non-GAAP measures and adjustments are issuer-specific. The SEC’s company submission APIs provide real-time filing history and do not require API keys, which makes this public-data architecture appropriate for the task. [1]

## Validation performed

| Check | Result |
| --- | --- |
| Python compilation of application, engine, and tests | Passed |
| Whitespace and patch integrity check (`git diff --check`) | Passed |
| Focused engine tests | Passed: 5 tests |
| Streamlit initial-state smoke test | Passed: 1 test |
| Full automated suite | Passed: **6 tests in 4.40 seconds** |
| Local Streamlit startup | Passed: health endpoint returned `ok`; root returned HTTP 200 |
| Live deployment visual review | App canvas remained blank in the unauthenticated audit browser; interactive issuer analysis was not performed |

No live SEC company run was performed because the application correctly requires a real contact email for the SEC User-Agent, and no such contact identity was provided for this audit. The code’s existing 0.12-second request interval, retry behavior, and user-agent handling were preserved.

## Findings that should be addressed next

### 1. Restore and verify the public deployment first

The current live-page observation is the highest operational priority. Open the Streamlit Cloud deployment logs and confirm that the deployed branch has both updated files from the same release. Then restart the app and test a known issuer from a normal browser session. The code already displays an explicit error if `app.py` and `sec_nongaap.py` are from incompatible releases, but this should be verified in the deployment logs rather than inferred from the blank canvas.

### 2. Add a small curated fixture library before expanding the extraction rules

The parser supports many formats, but the repository originally had no tests. Keep the new unit tests and add sanitized HTML/PDF fixtures for six to ten distinct issuer formats. The fixture set should include a normal HTML press release, a financial supplement, a text-heavy PDF, an image-only PDF, a company with multiple 8-Ks near the period end, a non-calendar fiscal year, and an issuer whose definition appears in narrative rather than a reconciliation table. Each fixture should assert the metric, GAAP endpoint, non-GAAP endpoint, adjustment items, definition evidence, source page, and tie-out status.

### 3. Make the result taxonomy more explicit on the landing state

The app has the desired information, but the initial instructions describe results as a list rather than a user journey. Replace the landing copy with three plain-language phases: **Discover measures**, **Review reconciliations**, and **Compare peers and definitions**. Add a one-line distinction between “Additional measures” and “Definitions & method.” The former identifies a measure mentioned in the exhibit; the latter requires language that appears to describe how the issuer defines or calculates it.

### 4. Introduce coverage and quality indicators at the metric level

The source audit is strong, but users must inspect several tabs to understand quality. Add compact badges for each metric/period: `Source found`, `Structured bridge`, `Line items parsed`, `Tie-out passed`, and `Definition evidence found`. This would turn the current audit trail into an immediate review queue. It should never silently suppress a failed tie-out or missing definition.

### 5. Add cache observability and a manual refresh control

The implemented caches are intentionally time-bounded at one hour. Streamlit supports TTL-based data caching and returns cached data as copies, which is a suitable model for public filings and derived DataFrames. [2] Add a sidebar caption such as “Results may be cached for up to one hour” and a “Refresh SEC data” button that clears the relevant cache entries. This makes the freshness trade-off visible to users without sacrificing the speed benefit.

### 6. Plan a durable ingestion architecture before growing the peer limit

The current interactive design is appropriate for a small peer set. Scaling it to many companies should not mean adding more concurrent browser requests. A future production architecture should use a queue-backed ingestion worker, a persistent database for raw document metadata and parsed results, and a per-document content hash. This would allow the user interface to return an existing analysis immediately, display ingestion progress, and reprocess only changed or newly filed documents. Public SEC submissions are updated throughout the day, so a durable data layer should record retrieval time and filing accession numbers. [1]

## Deployment instructions

The improvement set is ready to commit to the repository’s deployment branch. After the push completes, let Streamlit Cloud rebuild from that branch and then perform one manual end-to-end run using a real SEC contact email and a known issuer with a recent earnings release.

The recommended smoke-test sequence is: open the app; enter a valid contact email; search a liquid US issuer by ticker; load fiscal periods; analyze the latest one or two fiscal years; confirm that a source audit record exists; inspect one bridge; inspect the new definition-evidence tab; export the Excel workbook; and run a two-company peer comparison. Record any extraction variance as an issue with the SEC exhibit URL and expected source page.

## Basis, time, assumptions, and sources

**Basis.** The application treats the issuer’s 8-K and EX-99 exhibits as the only source for non-GAAP values, reconciliation items, measure mentions, and definition evidence. Periodic reports are used only for fiscal-calendar anchoring and event matching. Definition results are source text, not standardized metric definitions.

**Time.** The audit was performed on September 17, 2026. The repository head reviewed was `175b99e9b412475731a2f3758e2970503ac0b2d0`, dated August 18, 2026. Local tests were executed against the revised working copy on the audit date.

**Assumptions.** The implemented one-hour caching window is appropriate for interactive public-filing research, not for real-time filing surveillance. The 160 MB document-cache ceiling balances repeat-use performance with memory safety for a small peer set. The app should retain the user-entered real SEC contact email requirement.

**Sources and confidence.** High confidence in the repository findings and local validation. The live-site rendering observation is limited to the unauthenticated audit browser and should be confirmed in the deployment logs and a normal browser session. SEC API behavior and Streamlit caching semantics are based on the cited primary documentation.

**Compliance.** This is research and software analysis only, not personalized financial advice.

## References

[1]: https://www.sec.gov/search-filings/edgar-application-programming-interfaces "SEC EDGAR Application Programming Interfaces"
[2]: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.cache_data "Streamlit st.cache_data API reference"
[3]: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.cache_resource "Streamlit st.cache_resource API reference"
