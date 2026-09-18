# Live SEC Validation Entity Card

| Field | Apple Inc. |
| --- | --- |
| Legal / common name | Apple Inc. |
| Ticker / exchange | AAPL / Nasdaq |
| Listing status | U.S. public company |
| CIK | 0000320193 |
| Fiscal year end | September (to be confirmed against live SEC submissions metadata) |
| Reporting currency | U.S. dollars |
| Industry | Consumer electronics, software, and services |
| Validation purpose | Confirm live SEC retrieval, filing matching, resilience logging, and generated Excel workbook behavior. |

## Validation scope

The test uses the SEC contact identity supplied by the app owner only at runtime. It verifies retrieval from SEC-hosted sources, the matched earnings 8-K workflow, generated workbook sheets, and the absence or presence of request-resilience events. It is a software validation, not an investment analysis or a conclusion about Apple’s financial performance.

## Supplemental populated-workbook entity

| Field | Lattice Semiconductor Corporation |
| --- | --- |
| Ticker / exchange | LSCC / Nasdaq |
| Listing status | U.S. public company |
| CIK | 0000855658 |
| Validation purpose | Exercise the same live extraction and Excel path with a recent 8-K package that produces structured reconciliation and adjustment rows. |
