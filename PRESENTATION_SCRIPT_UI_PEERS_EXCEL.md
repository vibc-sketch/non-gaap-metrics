# Presentation Script: SEC Non-GAAP Metrics Explorer Improvements

**Audience:** Finance, research, and product stakeholders  
**Suggested duration:** 8–10 minutes  
**Purpose:** Explain the redesigned user experience, more reliable peer-company selection, stronger reconciliation quality controls, and presentation-ready Excel exports.

## Slide 1 — From Filing Retrieval to an Evidence-First Research Workspace

**On-screen message:**

> **SEC Non-GAAP Reconciliation & Peer Benchmarking**  
> Evidence-first analysis of earnings 8-K exhibits, reconciliations, and peer disclosures

**Speaker script:**

“This release moves the product beyond a filing-search utility and toward an evidence-first financial research workspace. The product now gives a user a structured path from issuer selection, to matched earnings 8-K exhibits, to GAAP-to-non-GAAP bridges, adjustment evidence, peer comparisons, and Excel output.

The central design principle is straightforward: the application should preserve what the issuer reported, identify what was parsed with confidence, and make every uncertain result visible rather than silently smoothing over a gap. The result is a more defensible workflow for reviewing non-GAAP disclosures.” [1]

**Transition:** “The first improvement is the user interface, because trusted financial analysis begins with readable information.”

## Slide 2 — BCG-Inspired Interface: Clear Hierarchy and Consistent Contrast

**On-screen message:**

| Design element | New treatment | User benefit |
|---|---|---|
| Product shell | Forest-green command bar and sidebar | A consistent enterprise identity across views |
| Main title | Full product title in the persistent top bar | Product context remains visible when navigation is collapsed |
| Workspace content | Light canvas and dark-ink text | Better readability for financial tables and source notes |
| Navigation | Wrapped, high-contrast tab controls | All analysis sections remain visible on constrained screens |
| Alerts and links | Dark text on pale amber; underlined blue SEC links | Warnings and source actions do not disappear in dark-mode environments |

**Speaker script:**

“We rebuilt the visual hierarchy around a disciplined, BCG-inspired palette. The forest-green top bar establishes the product identity. The full title, ‘SEC Non-GAAP Reconciliation & Peer Benchmarking,’ now lives in that persistent command bar. The workspace below uses a lighter, quieter surface for analysis.

More importantly, this is not only a cosmetic redesign. We explicitly force light data surfaces and dark text so the application does not inherit a browser or Streamlit dark-mode preference that makes warning copy, source links, or table labels unreadable. Active tabs use white text on forest green. Inactive tabs use dark text on white. Warning panels use dark-brown text on pale amber, and SEC links are darker blue and underlined.” [1]

**Transition:** “The interface makes evidence easier to consume. The next improvement makes peer selection easier to start.”

## Slide 3 — Peer Resolution: Start With How Users Actually Identify Companies

**On-screen message:**

> Enter peers by **ticker**, **company name**, **recognizable name fragment**, or **legal-name variant**.

| User entry | Example result |
|---|---|
| Ticker | `LSCC` → Lattice Semiconductor Corp. |
| Company name | `Lattice Semiconductor` → LSCC |
| Legal-name variant | `Lattice Semiconductor Corporation` → LSCC |
| List input | `AAPL MSFT AMD` or `Lattice Semiconductor, Marvell Technology` |

**Speaker script:**

“Before this release, peer benchmarking assumed that users entered valid ticker symbols. That is unnecessarily restrictive in a research workflow. Analysts often begin with company names, partial names, legal names, or a mixed list of names and symbols.

The peer input now accepts all of those forms. It supports commas, semicolons, and one company per line. It also accepts a space-separated all-caps ticker list. The resolver searches the SEC company universe, normalizes common legal suffixes such as ‘Corporation,’ ‘Corp.,’ ‘Inc.,’ and ‘Holdings,’ and then deduplicates companies using the SEC CIK identifier. This prevents the application from analyzing the same company twice when the user enters both its ticker and legal name.” [1]

**Transition:** “Once peers are resolved, the application needs to be equally careful with the reconciliation data it extracts.”

## Slide 4 — Reconciliation Quality: Preserve Evidence, Do Not Overstate Certainty

**On-screen message:**

> **Incomplete reconciliation detail — review required**
>
> The application found the GAAP and non-GAAP endpoints but could not verify every individual adjustment line.

**Speaker script:**

“This message is a quality-control feature, not an error message about the company. It means the application successfully identified the reported GAAP amount and the reported non-GAAP amount, but did not confidently extract every adjustment row from the source table.

In that case, the total adjustment shown in the bridge is simply the reported non-GAAP endpoint minus the reported GAAP endpoint. It is not represented as a verified sum of adjustment lines. The user is directed to the SEC exhibit and the Source audit tab so the original disclosure can be reviewed.

We also improved the parser itself. It now suppresses mixed-unit pairings, such as a dollar anchor incorrectly paired with a percentage endpoint from an adjacent table. It filters non-comparable rows from dollar bridges. It keeps issuer subtotals as audit evidence without double-counting them in individual adjustment totals. Finally, it distinguishes a true source-review issue from a case where an issuer subtotal confirms the endpoint difference.” [1]

**Transition:** “The practical effect is a smaller number of misleading bridges and a larger share of reconciliations that tie within the accepted rounding threshold.”

## Slide 5 — Parsing Results: Fewer Misleading Bridges, More Defensible Outputs

**On-screen message:**

| Live LSCC FY2026 validation result | Count |
|---|---:|
| Structured reconciliations retained | 35 |
| Ties within rounding | 22 |
| No parsed line-item detail | 8 |
| Needs source review | 5 |

**Speaker script:**

“We tested the updated logic on live Lattice Semiconductor earnings-release materials. The updated parser retained 35 structured reconciliations. Twenty-two tied within rounding. Eight had clear GAAP and non-GAAP endpoints but no reliable individual line-item detail. Five remained genuinely in need of source review.

This is the desired behavior. The product should not maximize the raw count of extracted bridges. It should maximize the count of bridges that are useful and defensible, while clearly labeling anything that requires analyst review.”

**Transition:** “That same design philosophy carries into the Excel deliverable.”

## Slide 6 — Professional Excel: From Raw Export to Review-Ready Reconciliation Bridges

**On-screen message:**

> **Reconciliation Bridges**  
> GAAP → individual adjustments → non-GAAP  
> Source-linked, issuer-labeled, and audit-ready

**Speaker script:**

“The Excel workbook has been redesigned so that the reconciliation output resembles the analytical hierarchy that finance teams expect to see in an earnings-release bridge.

The Reconciliation Bridges sheet now begins with a company title bar and a clear description of the source basis. Each period-and-metric bridge is presented as a discrete block. The GAAP starting point is emphasized, individual adjustment lines are indented, and the non-GAAP endpoint is emphasized with a final border.

The workbook retains the issuer’s exact labels and reported values. SEC-sourced values are blue, every source cell includes an audit comment, and the SEC exhibit is available through a clickable hyperlink. The sheet is set for landscape printing and fit-to-width review. Detailed machine-readable outputs remain in separate audit sheets, so the presentation layer stays clean without sacrificing traceability.” [1]

**Transition:** “The result is a more complete workflow, from search through review and deliverable creation.”

## Slide 7 — End-to-End User Workflow

**On-screen message:**

1. Search an issuer and load fiscal-period anchors.  
2. Analyze matched Item 2.02 earnings 8-K packages and EX-99 exhibits.  
3. Review bridges, parsing status, and source audit evidence.  
4. Add peers by name or ticker.  
5. Compare adjustment disclosures and trends.  
6. Export a review-ready, source-linked Excel workbook.

**Speaker script:**

“The improved workflow supports both fast discovery and analyst-grade review. A user begins with an issuer search, then the product identifies fiscal-period anchors and analyzes matched earnings 8-K exhibits. The user can inspect the reconciliation bridge, identify incomplete details, and open the source exhibit directly.

For benchmarking, the user can enter peers using the terminology they know: a ticker, a company name, or a legal-name variant. The Peer benchmark tab compares disclosed measures and adjustment categories, while the trend charts show disclosure presence over time without aggregating non-comparable dollar amounts across companies.

Finally, the Excel export provides a clear presentation layer for review and a source-linked audit layer for diligence.”

## Slide 8 — Closing: What Has Changed

**On-screen message:**

| Area | Outcome |
|---|---|
| User experience | A readable, stable enterprise workspace across desktop, mobile, and color-mode settings |
| Peer benchmarking | Flexible name-and-ticker lookup with SEC CIK-based deduplication |
| Parsing quality | Clear separation between tied bridges, incomplete detail, and source-review cases |
| Excel deliverable | Presentation-ready reconciliation bridges with source links and audit comments |

**Speaker script:**

“The central outcome of this release is stronger trust in the research workflow. The application is more readable, easier to use, more tolerant of natural peer-company inputs, more cautious when parsing evidence is incomplete, and more useful when the work needs to be shared in Excel.

The product now makes an explicit distinction between what the issuer reported, what the parser verified, and what an analyst should review. That distinction is what turns a data-extraction tool into a credible non-GAAP research platform.”

## References

[1]: https://github.com/vibc-sketch/non-gaap-metrics/commit/e9316c7 "GitHub commit e9316c7: Improve peer lookup and reconciliation outputs"
