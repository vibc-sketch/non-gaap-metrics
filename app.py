import io
import re
import time
import os
import math
from typing import Optional

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

SEC_HEADERS = {
    "User-Agent": os.getenv("SEC_USER_AGENT", "SEC Non-GAAP Metrics Explorer/1.0; contact: you@example.com"),
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov",
}
DATA_HEADERS = {**SEC_HEADERS, "Host": "data.sec.gov"}

st.set_page_config(page_title="SEC Non-GAAP Metrics Explorer", page_icon="📊", layout="wide")


def set_contact(email: str):
    if email.strip():
        ua = f"SEC Non-GAAP Metrics Explorer/1.0; {email.strip()}"
        SEC_HEADERS["User-Agent"] = ua
        DATA_HEADERS["User-Agent"] = ua


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def sec_get_json(url: str, data_api: bool = False):
    headers = DATA_HEADERS if data_api else SEC_HEADERS
    r = requests.get(url, headers=headers, timeout=45)
    r.raise_for_status()
    time.sleep(0.15)
    return r.json()


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def sec_get_text(url: str):
    r = requests.get(url, headers=SEC_HEADERS, timeout=60)
    r.raise_for_status()
    time.sleep(0.15)
    return r.text


@st.cache_data(ttl=30 * 24 * 3600, show_spinner=False)
def load_company_tickers():
    data = sec_get_json("https://www.sec.gov/files/company_tickers.json")
    rows = []
    for x in data.values():
        rows.append({"cik": int(x["cik_str"]), "ticker": str(x.get("ticker", "")).upper(), "name": x.get("title", "")})
    df = pd.DataFrame(rows)
    return df.sort_values(["name", "ticker"]).reset_index(drop=True)


@st.cache_data(ttl=30 * 24 * 3600, show_spinner=False)
def load_sic_codes():
    url = "https://www.sec.gov/search-filings/standard-industrial-classification-sic-code-list"
    html = sec_get_text(url)
    try:
        tables = pd.read_html(io.StringIO(html))
    except ValueError:
        return pd.DataFrame(columns=["sic", "industry_title"])
    if not tables:
        return pd.DataFrame(columns=["sic", "industry_title"])
    df = tables[0].copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    sic_col = next((c for c in df.columns if "sic" in c), None)
    ind_col = next((c for c in df.columns if "industry" in c and "title" in c), None)
    if not sic_col or not ind_col:
        return pd.DataFrame(columns=["sic", "industry_title"])
    out = df[[sic_col, ind_col]].rename(columns={sic_col: "sic", ind_col: "industry_title"})
    out["sic"] = out["sic"].astype(str).str.extract(r"(\d+)")[0]
    return out.dropna().drop_duplicates()


@st.cache_data(ttl=30 * 24 * 3600, show_spinner=False)
def browse_edgar_sic(sic_code: str):
    url = f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&sic={sic_code}&owner=exclude&count=100"
    html = sec_get_text(url)
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        m = re.search(r"CIK=(\d+)", href, re.I)
        txt = re.sub(r"\s+", " ", a.get_text(" ", strip=True))
        if m and txt:
            rows.append({"cik": int(m.group(1)), "name": txt})
    return pd.DataFrame(rows).drop_duplicates("cik") if rows else pd.DataFrame(columns=["cik", "name"])


@st.cache_data(ttl=30 * 24 * 3600, show_spinner=False)
def load_company_submissions(cik: int):
    return sec_get_json(f"https://data.sec.gov/submissions/CIK{int(cik):010d}.json", data_api=True)


def recent_filings(submissions: dict) -> pd.DataFrame:
    rec = submissions.get("filings", {}).get("recent", {}) or {}
    # Never assume all SEC arrays are present; construct a stable schema for the UI.
    if not isinstance(rec, dict) or not rec:
        return pd.DataFrame(columns=["form", "filingDate", "filed", "reportDate", "acceptanceDateTime", "act", "fileNumber", "filmNumber", "items", "coreg", "size", "isXBRL", "isInlineXBRL", "isXBRL", "isInlineXBRL", "accessionNumber", "primaryDocument", "primaryDocDescription", "form", "fileNumber", "filmNumber", "items", "coreg"])
    df = pd.DataFrame(rec)
    # Keep canonical column names and add safe fallbacks for missing metadata.
    for c in ["form", "accessionNumber", "primaryDocument", "reportDate", "filed", "fy", "fp", "frame"]:
        if c not in df.columns:
            df[c] = pd.NA
    for c in ["filed", "reportDate"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def fiscal_year_from_date(report_date, fy_end: str):
    if pd.isna(report_date):
        return None
    try:
        d = pd.Timestamp(report_date)
        if fy_end and re.fullmatch(r"\d{4}", str(fy_end)):
            m = int(str(fy_end)[:2]); day = int(str(fy_end)[2:])
            # fiscal year is the year in which the fiscal-year-end occurs.
            return d.year if (d.month, d.day) <= (m, day) else d.year + 1
        return d.year
    except Exception:
        return None


def fiscal_period_from_dates(report_date, fy_end: str):
    if pd.isna(report_date) or not fy_end or not re.fullmatch(r"\d{4}", str(fy_end)):
        return "N/A"
    try:
        d = pd.Timestamp(report_date)
        m = int(str(fy_end)[:2]); day = int(str(fy_end)[2:])
        end = pd.Timestamp(year=d.year, month=m, day=day)
        if d > end:
            end = end.replace(year=end.year + 1)
        months_back = (end.year - d.year) * 12 + (end.month - d.month)
        # Quarter ending closest to 3/6/9/12 months before fiscal-year-end.
        q = min(4, max(1, 4 - int(round(months_back / 3.0))))
        return f"Q{q}"
    except Exception:
        return "N/A"


def normalized_period(row, fy_end: str):
    raw_fy = row.get("fy")
    raw_fp = row.get("fp")
    fy = None
    if pd.notna(raw_fy):
        try: fy = int(raw_fy)
        except Exception: pass
    if not fy:
        fy = fiscal_year_from_date(row.get("reportDate"), fy_end)
    fp = str(raw_fp).upper() if pd.notna(raw_fp) else ""
    if fp not in {"Q1", "Q2", "Q3", "Q4", "FY"}:
        fp = "FY" if str(row.get("form", "")).startswith("10-K") else fiscal_period_from_dates(row.get("reportDate"), fy_end)
    return fy, fp


def filing_url(cik: int, accession: str, primary_doc: str):
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession.replace('-', '')}/{primary_doc}"


def index_url(cik: int, accession: str):
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession.replace('-', '')}/{accession}-index.htm"


def extract_exhibits_from_filing(cik: int, accession: str):
    html = sec_get_text(index_url(cik, accession))
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for tr in soup.find_all("tr"):
        a = tr.find("a", href=True)
        cells = tr.find_all(["td", "th"])
        if not a or len(cells) < 2:
            continue
        desc = " | ".join(re.sub(r"\s+", " ", c.get_text(" ", strip=True)) for c in cells)
        href = a.get("href", "")
        if href.startswith("/"):
            url = "https://www.sec.gov" + href
        elif href.startswith("http"):
            url = href
        else:
            url = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession.replace('-', '')}/{href}"
        doc = a.get_text(" ", strip=True)
        if re.search(r"EX-99\.[0-9]+|99\.[0-9]+|press release|earnings|financial results|results of operations", desc, re.I):
            rows.append({"document": doc, "description": desc, "url": url})
    # De-duplicate while preserving order.
    seen = set(); out = []
    for r in rows:
        if r["url"] not in seen:
            seen.add(r["url"]); out.append(r)
    return out


def clean_text(html):
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript"]): tag.decompose()
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))


def metric_family(text: str):
    t = text.lower()
    rules = [
        ("Adjusted EBITDA", ["adjusted ebitda", "adjusted earnings before interest"]),
        ("Adjusted EPS", ["adjusted eps", "non-gaap eps", "adjusted earnings per share"]),
        ("Adjusted Net Income", ["adjusted net income", "non-gaap net income"]),
        ("Free Cash Flow", ["free cash flow", "fcf"]),
        ("Adjusted Operating Income", ["adjusted operating income"]),
        ("Organic Revenue", ["organic revenue", "organic sales"]),
        ("Constant Currency", ["constant currency"]),
        ("Net Debt / Leverage", ["net debt", "leverage ratio"]),
        ("Non-GAAP", ["non-gaap", "non gaap"]),
    ]
    for label, needles in rules:
        if any(x in t for x in needles): return label
    return "Other Non-GAAP"


def extract_non_gaap(html: str):
    soup = BeautifulSoup(html, "html.parser")
    out = []
    seen = set()

    def add(family, text, source_type, label=""):
        text = re.sub(r"\s+", " ", text).strip()
        key = text.lower()
        if not text or key in seen or len(text) > 1200:
            return
        seen.add(key)
        out.append({"metric_family": family or metric_family(text), "metric_label": label, "disclosure": text, "source_type": source_type})

    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            vals = [re.sub(r"\s+", " ", x.get_text(" ", strip=True)) for x in tr.find_all(["th", "td"])]
            if not vals: continue
            joined = " | ".join(vals)
            if re.search(r"adjusted|non[- ]gaap|free cash flow|ebitda|organic|constant currency|net debt|leverage|pro forma|underlying|normalized", joined, re.I):
                label = vals[0][:160]
                add(metric_family(joined), joined, "table", label)

    text = clean_text(html)
    # Pull a useful window around each key metric phrase rather than a brittle value regex.
    patterns = [
        r"(?:adjusted ebitda|adjusted eps|adjusted earnings per share|adjusted net income|free cash flow|non[- ]gaap|organic revenue|constant currency|net debt|leverage)[^.;]{0,320}",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.I):
            snippet = m.group(0)
            add(metric_family(snippet), snippet, "narrative", snippet[:120])
    return out[:500]


def extract_filing_data(cik, filings, fy_end, selected_years):
    if filings.empty:
        return pd.DataFrame()
    forms = {"8-K", "8-K/A", "10-K", "10-Q", "10-Q/A", "S-1", "S-1/A"}
    f = filings[filings["form"].isin(forms)].copy()
    f["norm_fy"] = f.apply(lambda r: normalized_period(r, fy_end)[0], axis=1)
    f["norm_fp"] = f.apply(lambda r: normalized_period(r, fy_end)[1], axis=1)
    f = f[f["norm_fy"].isin(selected_years)]
    rows = []
    for _, r in f.sort_values(["norm_fy", "norm_fp", "filed"], ascending=[False, True, False]).iterrows():
        acc = r.get("accessionNumber"); doc = r.get("primaryDocument")
        if pd.isna(acc) or pd.isna(doc): continue
        form = str(r.get("form", ""))
        primary_url = filing_url(cik, acc, doc)
        try:
            exhibits = extract_exhibits_from_filing(cik, acc)
        except Exception:
            exhibits = []
        # 10-K/10-Q often contain MD&A but earnings-release data is normally in 8-K exhibits.
        candidates = exhibits
        if not candidates and form in {"10-K", "10-Q", "10-Q/A", "S-1", "S-1/A"}:
            candidates = [{"document": str(doc), "description": "Primary filing", "url": primary_url}]
        # For 8-K, prioritize earnings/financial-results exhibits.
        candidates = sorted(candidates, key=lambda x: (0 if re.search(r"EX-99|press release|earnings|financial results", x["description"], re.I) else 1, x["document"]))
        for ex in candidates[:8]:
            try:
                html = sec_get_text(ex["url"])
                hits = extract_non_gaap(html)
            except Exception:
                hits = []
            for hit in hits:
                rows.append({
                    "fiscal_year": int(r["norm_fy"]) if pd.notna(r["norm_fy"]) else None,
                    "fiscal_period": r["norm_fp"],
                    "filing_form": form,
                    "filing_date": pd.Timestamp(r["filed"]).date().isoformat() if pd.notna(r["filed"]) else "",
                    "report_date": pd.Timestamp(r["reportDate"]).date().isoformat() if pd.notna(r["reportDate"]) else "",
                    "metric_family": hit["metric_family"],
                    "metric_label": hit["metric_label"],
                    "disclosure": hit["disclosure"],
                    "source_type": hit["source_type"],
                    "source_url": ex["url"],
                    "filing_url": primary_url,
                })
    return pd.DataFrame(rows)


def search_companies(query="", sic="", industry="", limit=50):
    tickers = load_company_tickers()
    q = (query or "").strip().upper()
    if q:
        df = tickers[tickers["ticker"].str.contains(re.escape(q), na=False) | tickers["name"].str.upper().str.contains(re.escape(q), na=False)].copy()
    else:
        df = tickers.copy()
    if sic.strip().isdigit():
        sic_df = browse_edgar_sic(sic.strip())
        df = df[df["cik"].isin(sic_df["cik"])] if not sic_df.empty else df.iloc[0:0]
    elif industry.strip():
        sic_df = load_sic_codes()
        matches = sic_df[sic_df["industry_title"].str.contains(re.escape(industry.strip()), case=False, na=False)]
        frames = []
        for code in matches["sic"].head(20).tolist():
            try: frames.append(browse_edgar_sic(code))
            except Exception: pass
        if frames:
            ids = pd.concat(frames)["cik"].drop_duplicates()
            df = df[df["cik"].isin(ids)]
        else:
            df = df.iloc[0:0]
    return df.head(limit).reset_index(drop=True)


def company_record(cik):
    sub = load_company_submissions(cik)
    return {
        "cik": int(cik), "name": sub.get("name", ""), "ticker": ", ".join(sub.get("tickers", []) or []),
        "sic": str(sub.get("sic", "")), "sic_description": sub.get("sicDescription", ""),
        "fiscal_year_end": str(sub.get("fiscalYearEnd", "")),
    }

# UI
st.title("SEC Non-GAAP Metrics Explorer")
st.caption("Search public issuers by name, ticker, SIC, or industry; pull quarterly non-GAAP disclosures and normalize them to the issuer's fiscal calendar.")

with st.sidebar:
    st.header("Issuer search")
    lookup = st.text_input("Name or ticker", placeholder="e.g. LSCC or Lattice Semiconductor")
    sic = st.text_input("SIC code (optional)", placeholder="e.g. 3674")
    industry = st.text_input("Industry keyword (optional)", placeholder="e.g. semiconductors")
    contact = st.text_input("SEC contact email", value=os.getenv("SEC_CONTACT_EMAIL", ""), placeholder="you@example.com")
    if st.button("Find issuers", type="primary", use_container_width=True):
        if not any(x.strip() for x in [lookup, sic, industry]):
            st.warning("Enter a name/ticker, SIC code, or industry keyword.")
        else:
            try:
                set_contact(contact)
                with st.spinner("Searching SEC issuer universe…"):
                    st.session_state["issuer_search"] = search_companies(lookup, sic, industry)
            except Exception as e:
                st.error(f"Search failed: {e}")
    st.markdown("**Sources**")
    st.write("8-K / 8-K-A, 10-K / 10-Q / amendments, S-1 / S-1-A exhibits")

results = st.session_state.get("issuer_search", pd.DataFrame())
if not results.empty:
    st.subheader("Issuer matches")
    options = {i: f"{r['name']} ({r['ticker'] or 'no ticker'}) — CIK {r['cik']}" for i, r in results.iterrows()}
    choice = st.selectbox("Select an issuer", list(options.keys()), format_func=lambda i: options[i])
    if st.button("Load issuer filings", use_container_width=True):
        st.session_state["selected_cik"] = int(results.loc[choice, "cik"])

cik = st.session_state.get("selected_cik")
if cik:
    try:
        company = company_record(cik)
        filings = recent_filings(load_company_submissions(cik))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Company", company["name"][:30])
        c2.metric("Ticker", company["ticker"] or "—")
        c3.metric("SIC", company["sic"] or "—")
        c4.metric("Fiscal year end", company["fiscal_year_end"] or "—")

        # Make a normalized metadata table before any .unique()/astype operation.
        filings["norm_fy"] = filings.apply(lambda r: normalized_period(r, company["fiscal_year_end"])[0], axis=1)
        filings["norm_fp"] = filings.apply(lambda r: normalized_period(r, company["fiscal_year_end"])[1], axis=1)
        years = sorted({int(x) for x in filings["norm_fy"].dropna().tolist() if int(x) > 0}, reverse=True)
        if not years:
            st.warning("SEC did not return fiscal-year metadata for this issuer. The app can still use report dates, but no fiscal-year filter is available.")
        default_years = years[:2]
        selected_years = st.multiselect("Fiscal years", years, default=default_years, max_selections=2)
        include_s1 = st.checkbox("Include S-1 / S-1-A", value=True)
        forms = {"8-K", "8-K/A", "10-K", "10-Q", "10-Q/A"} | ({"S-1", "S-1/A"} if include_s1 else set())
        filt = filings[filings["form"].isin(forms)].copy()
        if selected_years:
            filt = filt[filt["norm_fy"].isin(selected_years)]
        st.write(f"**{len(filt)}** relevant filings in the selected fiscal years.")
        if not filt.empty:
            st.dataframe(filt[["form", "filed", "reportDate", "norm_fy", "norm_fp", "accessionNumber", "primaryDocument"]], use_container_width=True, height=260)

        if st.button("Pull non-GAAP analysis", type="primary", use_container_width=True):
            if not selected_years:
                st.warning("Select at least one fiscal year.")
            else:
                with st.spinner("Reading SEC filings and earnings-release exhibits…"):
                    data = extract_filing_data(cik, filings, company["fiscal_year_end"], selected_years)
                st.session_state["metrics"] = data

        data = st.session_state.get("metrics", pd.DataFrame())
        if not data.empty:
            st.subheader("Quarter-by-quarter non-GAAP analysis")
            # Coverage grid
            cov = data.groupby(["fiscal_year", "fiscal_period"]).size().reset_index(name="disclosures")
            cov["period"] = cov["fiscal_year"].astype(str) + " " + cov["fiscal_period"]
            st.dataframe(cov[["period", "disclosures"]], use_container_width=True, hide_index=True)

            # Executive summary by metric family and quarter.
            summary = data.groupby(["metric_family", "fiscal_year", "fiscal_period"]).size().reset_index(name="disclosure_count")
            st.markdown("### Metric coverage")
            st.dataframe(summary.sort_values(["fiscal_year", "fiscal_period", "metric_family"], ascending=[False, True, True]), use_container_width=True, hide_index=True)

            st.markdown("### Evidence")
            shown = data.sort_values(["fiscal_year", "fiscal_period", "metric_family", "filing_date"], ascending=[False, True, True, False]).copy()
            st.dataframe(shown[["fiscal_year", "fiscal_period", "metric_family", "metric_label", "disclosure", "filing_form", "source_type", "source_url", "filing_url"]], use_container_width=True, height=560,
                         column_config={"source_url": st.column_config.LinkColumn("Source"), "filing_url": st.column_config.LinkColumn("Filing")}, hide_index=True)
            st.download_button("Download normalized CSV", data.to_csv(index=False).encode("utf-8"), file_name=f"{company['ticker'] or cik}_non_gaap_analysis.csv", mime="text/csv")
        elif selected_years:
            st.info("No non-GAAP disclosures were extracted yet. Click 'Pull non-GAAP analysis'.")
    except requests.HTTPError as e:
        st.error(f"SEC request failed: {e}. Add a real SEC contact email and retry.")
    except Exception as e:
        st.exception(e)

st.markdown("---")
st.caption("Non-GAAP extraction is evidence-preserving and heuristic. Always review the linked SEC filing/exhibit before relying on a metric for investment, valuation, or reporting.")
