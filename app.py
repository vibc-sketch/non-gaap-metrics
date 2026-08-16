import io
import re
import time
import json
import math
import os
from datetime import datetime, date
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

SEC_USER_AGENT = os.getenv("SEC_USER_AGENT", "SEC Non-GAAP Metrics Explorer/1.0; contact: your-email@example.com")
SEC_HEADERS = {
    "User-Agent": SEC_USER_AGENT,
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov",
}
DATA_HEADERS = {**SEC_HEADERS, "Host": "data.sec.gov"}

st.set_page_config(page_title="SEC Non-GAAP Metrics Explorer", page_icon="📊", layout="wide")

@st.cache_data(ttl=24*3600, show_spinner=False)
def sec_get_json(url: str, data_api: bool=False):
    headers = DATA_HEADERS if data_api else SEC_HEADERS
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    time.sleep(0.12)
    return r.json()

@st.cache_data(ttl=7*24*3600, show_spinner=False)
def sec_get_text(url: str):
    r = requests.get(url, headers=SEC_HEADERS, timeout=45)
    r.raise_for_status()
    time.sleep(0.12)
    return r.text

@st.cache_data(ttl=30*24*3600, show_spinner=False)
def load_sic_codes():
    url = "https://www.sec.gov/search-filings/standard-industrial-classification-sic-code-list"
    html = sec_get_text(url)
    tables = pd.read_html(io.StringIO(html))
    if not tables:
        return pd.DataFrame(columns=["sic", "industry_title"])
    df = tables[0].copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    # SEC page currently exposes columns SIC Code / Office / Industry Title.
    sic_col = next((c for c in df.columns if "sic" in c), None)
    ind_col = next((c for c in df.columns if "industry" in c and "title" in c), None)
    if not sic_col or not ind_col:
        return pd.DataFrame(columns=["sic", "industry_title"])
    out = df[[sic_col, ind_col]].rename(columns={sic_col:"sic", ind_col:"industry_title"})
    out["sic"] = out["sic"].astype(str).str.extract(r"(\d+)")[0]
    out["industry_title"] = out["industry_title"].astype(str)
    return out.dropna().drop_duplicates()


def browse_edgar_sic(sic_code: str):
    url = f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&sic={sic_code}&owner=exclude&count=100"
    html = sec_get_text(url)
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for a in soup.find_all("a", href=True):
        txt = re.sub(r"\s+", " ", a.get_text(" ", strip=True))
        m = re.search(r"CIK=(\d+)", a.get("href", ""), re.I)
        if m and txt:
            rows.append({"cik": int(m.group(1)), "name": txt})
    return pd.DataFrame(rows).drop_duplicates("cik") if rows else pd.DataFrame(columns=["cik", "name"])


@st.cache_data(ttl=30*24*3600, show_spinner=False)
def load_company_tickers():
    data = sec_get_json("https://www.sec.gov/files/company_tickers.json")
    rows = []
    for _, x in data.items():
        rows.append({
            "cik": int(x["cik_str"]),
            "ticker": x.get("ticker", ""),
            "name": x.get("title", ""),
        })
    df = pd.DataFrame(rows)
    df["cik"] = df["cik"].astype(int)
    df["ticker"] = df["ticker"].fillna("").str.upper()
    return df

@st.cache_data(ttl=30*24*3600, show_spinner=False)
def load_company_submissions(cik: int):
    cik10 = f"{int(cik):010d}"
    return sec_get_json(f"https://data.sec.gov/submissions/CIK{cik10}.json", data_api=True)

@st.cache_data(ttl=30*24*3600, show_spinner=False)
def load_companyfacts(cik: int):
    cik10 = f"{int(cik):010d}"
    return sec_get_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik10}.json", data_api=True)


def search_companies(query: str, sic: Optional[str]=None, industry: Optional[str]=None, limit: int=50):
    tickers = load_company_tickers().copy()
    q = (query or "").strip().upper()

    # Name/ticker search is fast because SEC publishes a compact issuer index.
    if q:
        df = tickers[
            tickers["ticker"].str.contains(re.escape(q), na=False)
            | tickers["name"].str.upper().str.contains(re.escape(q), na=False)
        ].copy()
    else:
        df = tickers.copy()

    sic_codes = set()
    if sic and str(sic).strip().isdigit():
        sic_codes.add(str(sic).strip())
    if industry and not (sic and str(sic).strip().isdigit()):
        sic_df = load_sic_codes()
        needle = industry.strip().upper()
        matches = sic_df[sic_df["industry_title"].str.upper().str.contains(re.escape(needle), na=False)]
        sic_codes.update(matches["sic"].head(25).tolist())

    if sic_codes:
        # Query the SEC's SIC-filtered company index and intersect with name/ticker matches when present.
        candidates = []
        for code in sorted(sic_codes):
            try:
                candidates.append(browse_edgar_sic(code))
            except Exception:
                continue
        if candidates:
            sic_df = pd.concat(candidates, ignore_index=True).drop_duplicates("cik")
            df = df[df["cik"].isin(sic_df["cik"])] if q else tickers[tickers["cik"].isin(sic_df["cik"])]

    return df.head(limit)


def recent_filings(submissions: dict) -> pd.DataFrame:
    rec = submissions.get("filings", {}).get("recent", {})
    if not rec:
        return pd.DataFrame()
    df = pd.DataFrame(rec)
    for c in ["filed", "reportDate"]:
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def fiscal_quarter(fy, fp, report_date=None, fy_end=None):
    fp = str(fp or "").upper()
    if fp in {"FY", "Q1", "Q2", "Q3", "Q4"}:
        return fp
    if report_date and fy_end:
        try:
            rd = pd.Timestamp(report_date)
            end = pd.Timestamp(fy_end)
            # Approximate fallback for filings with missing FP; anchored to fiscal year end.
            months = ((end.month - rd.month) % 12)
            q = 4 - math.floor(months / 3)
            return f"Q{max(1, min(4, q))}"
        except Exception:
            pass
    return "N/A"


def filing_url(cik: int, accession: str, primary_doc: str):
    cik_path = str(int(cik))
    acc_nodash = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{cik_path}/{acc_nodash}/{primary_doc}"


def index_url(cik: int, accession: str):
    acc_nodash = accession.replace("-", "")
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc_nodash}/{accession}-index.htm"


def extract_exhibits_from_filing(index_url_: str, accession: str, cik: int):
    html = sec_get_text(index_url_)
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    # SEC filing index tables typically expose description, document, type, size.
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) < 2:
            continue
        text = [re.sub(r"\s+", " ", c.get_text(" ", strip=True)) for c in cells]
        a = tr.find("a", href=True)
        doc = a.get_text(" ", strip=True) if a else ""
        href = a["href"] if a else ""
        joined = " | ".join(text)
        if re.search(r"EX-99\.[0-9]+|PRESS RELEASE|EARNINGS|RESULTS", joined, re.I):
            if doc:
                if href.startswith("/"):
                    u = "https://www.sec.gov" + href
                elif href.startswith("http"):
                    u = href
                else:
                    acc_nodash = accession.replace("-", "")
                    u = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc_nodash}/{href}"
                rows.append({"description": joined, "document": doc, "url": u})
    return rows


def clean_text(html: str):
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))


def parse_number(raw: str):
    s = raw.replace(",", "").replace(" ", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if s.startswith("$"):
        s = s[1:]
    mult = 1.0
    if s.lower().endswith("m"):
        mult = 1e6; s = s[:-1]
    elif s.lower().endswith("b"):
        mult = 1e9; s = s[:-1]
    elif s.lower().endswith("k"):
        mult = 1e3; s = s[:-1]
    try:
        v = float(s) * mult
        return -v if neg else v
    except Exception:
        return None


def extract_non_gaap_from_html(html: str):
    soup = BeautifulSoup(html, "html.parser")
    hits = []
    # Table-first extraction: preserve row semantics when possible.
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            vals = [re.sub(r"\s+", " ", x.get_text(" ", strip=True)) for x in tr.find_all(["th", "td"])]
            if vals:
                rows.append(vals)
        if not rows:
            continue
        for row in rows:
            joined = " | ".join(row)
            if re.search(r"adjusted|non-gaap|non gaap|organic|free cash flow|ebitda|ebit|constant currency|pro forma|underlying|normalized|net debt|leverage", joined, re.I):
                hits.append({"metric_context": joined, "source_type": "table"})
    # Narrative metric/value pairs.
    text = clean_text(html)
    patterns = [
        r"(?P<label>(?:adjusted|non[- ]gaap|organic|free cash flow|adjusted EBITDA|EBITDA|adjusted EPS|adjusted earnings per share)[^.;:]{0,80}?)\s*(?:was|were|of|:|at)\s*(?P<value>\$?\(?\d[\d,\.]*\)?%?)",
        r"(?P<label>(?:adjusted|non[- ]gaap|organic|free cash flow|EBITDA)[^.;:]{0,80}?)\s*(?P<value>\$?\(?\d[\d,\.]*\)?%?)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.I):
            hits.append({"metric_context": f"{m.group('label').strip()}: {m.group('value')}", "source_type": "narrative"})
    # Deduplicate and rank likely metric-bearing snippets.
    seen = set(); out = []
    for h in hits:
        k = h["metric_context"][:500].lower()
        if k in seen: continue
        seen.add(k)
        out.append(h)
    return out[:250]


def extract_filing_data(cik: int, filings: pd.DataFrame, fy_end: str, years_back: int=2):
    if filings.empty:
        return pd.DataFrame()
    forms = {"8-K", "8-K/A", "10-K", "10-Q", "10-Q/A", "S-1", "S-1/A"}
    f = filings[filings["form"].isin(forms)].copy()
    f = f[f["filed"] >= (pd.Timestamp.today() - pd.DateOffset(years=years_back+1))]
    rows = []
    for _, r in f.sort_values("filed", ascending=False).iterrows():
        form = str(r.get("form", ""))
        acc = r.get("accessionNumber")
        doc = r.get("primaryDocument")
        report_date = r.get("reportDate")
        fy = r.get("fy")
        fp = r.get("fp")
        q = fiscal_quarter(fy, fp, report_date, fy_end)
        # For 8-K/S-1, identify relevant exhibits; for periodic reports still inspect exhibits for press releases.
        url = filing_url(cik, acc, doc)
        try:
            exs = extract_exhibits_from_filing(index_url(cik, acc), acc, cik)
        except Exception:
            exs = []
        candidates = exs[:]
        if not candidates and form in {"8-K", "8-K/A", "10-K", "10-Q", "10-Q/A", "S-1", "S-1/A"}:
            candidates = [{"description": "Primary filing", "document": doc, "url": url}]
        for ex in candidates:
            if not re.search(r"EX-99|PRESS RELEASE|EARNINGS|RESULTS|Primary filing|S-1", ex.get("description", ""), re.I):
                continue
            try:
                h = sec_get_text(ex["url"])
                snippets = extract_non_gaap_from_html(h)
            except Exception:
                snippets = []
            for s in snippets:
                rows.append({
                    "fiscal_year": int(fy) if pd.notna(fy) and str(fy).isdigit() else None,
                    "fiscal_period": q,
                    "fiscal_end": fy_end,
                    "filing_form": form,
                    "filing_date": pd.Timestamp(r["filed"]).date().isoformat() if pd.notna(r["filed"]) else "",
                    "report_date": pd.Timestamp(report_date).date().isoformat() if pd.notna(report_date) else "",
                    "accession": acc,
                    "metric_context": s["metric_context"],
                    "source_type": s["source_type"],
                    "source_url": ex["url"],
                    "filing_url": url,
                })
    return pd.DataFrame(rows)


def build_company_record(cik: int):
    sub = load_company_submissions(cik)
    return {
        "cik": cik,
        "name": sub.get("name", ""),
        "ticker": ", ".join(sub.get("tickers", []) or []),
        "sic": str(sub.get("sic", "")),
        "sic_description": sub.get("sicDescription", ""),
        "fiscal_year_end": sub.get("fiscalYearEnd", ""),
    }

st.title("SEC Non-GAAP Metrics Explorer")
st.caption("Pulls SEC filing data and exhibits, then organizes non-GAAP disclosures by issuer fiscal year and quarter.")

with st.sidebar:
    st.header("Issuer")
    lookup = st.text_input("Name or ticker", placeholder="e.g., Microsoft or MSFT")
    sic = st.text_input("SIC code (optional)", placeholder="e.g., 7372")
    industry = st.text_input("Industry keyword (optional)", placeholder="e.g., software")
    contact_default = os.getenv("SEC_CONTACT_EMAIL", "")
    contact = st.text_input("SEC contact email (recommended)", value=contact_default, placeholder="you@company.com")
    search_clicked = st.button("Find issuers", type="primary", use_container_width=True)
    st.markdown("**Supported sources**")
    st.write("8-K / 8-K-A, 10-K / 10-Q exhibits, S-1 / S-1-A")
    st.markdown("**Normalization**")
    st.write("Issuer fiscal year end → fiscal quarter; not calendar quarter.")

if search_clicked:
    if not lookup.strip() and not sic.strip() and not industry.strip():
        st.warning("Enter a company name/ticker, SIC code, or industry keyword.")
    else:
        if contact.strip():
            SEC_HEADERS["User-Agent"] = f"SEC Non-GAAP Metrics Explorer/1.0; {contact.strip()}"
            DATA_HEADERS["User-Agent"] = SEC_HEADERS["User-Agent"]
        with st.spinner("Searching SEC issuer universe…"):
            try:
                res = search_companies(lookup, sic=sic, industry=industry)
                st.session_state["issuer_search"] = res
            except requests.HTTPError as e:
                st.error(f"SEC request failed ({getattr(e.response, 'status_code', 'HTTP error')}). Check your SEC contact email and try again.")
            except Exception as e:
                st.error(f"Search failed: {e}")

results = st.session_state.get("issuer_search", pd.DataFrame())
if not results.empty:
    st.subheader("Issuer matches")
    display = results.rename(columns={"name":"Company", "ticker":"Ticker", "cik":"CIK"})[["Company","Ticker","CIK"]]
    choice = st.selectbox("Select an issuer", display.index, format_func=lambda i: f"{display.loc[i,'Company']} ({display.loc[i,'Ticker'] or 'no ticker'}) — CIK {display.loc[i,'CIK']}")
    if st.button("Load issuer filings", use_container_width=True):
        st.session_state["selected_cik"] = int(display.loc[choice, "CIK"])

cik = st.session_state.get("selected_cik")
if cik:
    try:
        company = build_company_record(cik)
        sub = load_company_submissions(cik)
        filings = recent_filings(sub)
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Company", company["name"][:28])
        c2.metric("Ticker", company["ticker"] or "—")
        c3.metric("SIC", company["sic"] or "—")
        c4.metric("Fiscal year end", company["fiscal_year_end"] or "—")

        years = sorted([int(x) for x in filings["fy"].dropna().astype(int).unique() if int(x) > 0], reverse=True)
        default_fy = years[:2] if len(years) >= 2 else years
        selected_years = st.multiselect("Fiscal years", years, default=default_fy[:2], max_selections=2)
        include_s1 = st.checkbox("Include S-1 / S-1-A offering disclosures", value=True)
        forms_filter = ["8-K", "8-K/A", "10-K", "10-Q", "10-Q/A"] + (["S-1", "S-1/A"] if include_s1 else [])
        if selected_years:
            filt = filings[filings["form"].isin(forms_filter)].copy()
            filt = filt[filt["fy"].fillna(0).astype(int).isin(selected_years)]
            st.write(f"Found **{len(filt)}** relevant filings for fiscal years **{', '.join(map(str, selected_years))}**.")
            if st.button("Extract non-GAAP metrics", type="primary"):
                with st.spinner("Reading filings and exhibits…"):
                    data = extract_filing_data(cik, filt, company["fiscal_year_end"], years_back=2)
                st.session_state["metrics"] = data

        data = st.session_state.get("metrics", pd.DataFrame())
        if not data.empty:
            st.subheader("Normalized non-GAAP disclosure set")
            st.caption("Each row is an extracted disclosure snippet; values are kept with their original units/context to avoid inventing conversions.")
            # Prefer fiscal-year + quarter ordering.
            data = data.sort_values(["fiscal_year","fiscal_period","filing_date"], ascending=[False, True, False], na_position="last")
            st.dataframe(data, use_container_width=True, height=520, column_config={
                "source_url": st.column_config.LinkColumn("Source"),
                "filing_url": st.column_config.LinkColumn("Filing"),
            })
            csv = data.to_csv(index=False).encode("utf-8")
            st.download_button("Download CSV", csv, file_name=f"{company['ticker'] or cik}_nongaap_metrics.csv", mime="text/csv")

            st.markdown("### Quarter coverage")
            coverage = data.drop_duplicates(["fiscal_year","fiscal_period"])[["fiscal_year","fiscal_period"]].copy()
            if not coverage.empty:
                coverage["quarter"] = coverage["fiscal_year"].astype(str) + " " + coverage["fiscal_period"].astype(str)
                st.write(" | ".join(coverage["quarter"].tolist()))
        else:
            st.info("Select up to two fiscal years and run extraction. The app will use SEC filing metadata to keep quarters aligned to the issuer’s fiscal calendar.")
    except requests.HTTPError as e:
        st.error(f"SEC request failed: {e}. Check your connection and try again.")
    except Exception as e:
        st.exception(e)

st.markdown("---")
st.caption("Prototype note: SEC filings are public-source documents. Non-GAAP extraction is heuristic and should be reviewed against the linked filing exhibit before investment or reporting use.")
