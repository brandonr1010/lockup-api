import os, io, json, math, time, requests, logging, re
from datetime import datetime, timedelta, date
from supabase import create_client

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

SUPABASE_URL  = os.environ["SUPABASE_URL"]
SUPABASE_KEY  = os.environ["SUPABASE_SERVICE_KEY"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
BUCKET        = "workbook"
FILE_NAME     = "Lockup_automation2.xlsm"
MAX_PER_RUN   = 5

SEC_HEADERS = {"User-Agent": "Brandon Ross brandonr1010@gmail.com"}

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def excel_round(val, decimals=0):
    factor = 10 ** decimals
    return math.floor(val * factor + 0.5) / factor

def calc_score(insider, float_pct, ev_sales, ev_ebitda, d, e):
    A = excel_round(min(insider / max(float_pct, 0.01), 5) / 5 * 30, 1)
    B = excel_round(min(insider / 100, 1) * 25, 1)
    # Valuation Risk (0-20): no earnings = max risk (20). NM means risk, not zero.
    def _is_nm(x):
        try:
            float(x); return False
        except (ValueError, TypeError):
            return True
    _s_nm = _is_nm(ev_sales)
    _e_nm = _is_nm(ev_ebitda)
    if _s_nm and _e_nm:
        cS, cE = 20.0, 0.0
    else:
        cS = 10.0 if _s_nm else min(float(ev_sales) / 5 * 9, 9)
        cE = 10.0 if _e_nm else min(max(float(ev_ebitda) - 5, 0) / 25 * 9, 9)
    C   = excel_round(cS + cE, 1)
    raw = excel_round(A + B + C + d + e, 1)
    score = int(excel_round(max(0, min(100, raw)), 0))
    tier  = "High" if score>=75 else "Medium" if score>=50 else "Low" if score>=25 else "Minimal"
    return score, tier

def download_workbook():
    return supabase.storage.from_(BUCKET).download(FILE_NAME)

def upload_workbook(file_bytes):
    supabase.storage.from_(BUCKET).update(
        FILE_NAME, file_bytes,
        {"content-type": "application/vnd.ms-excel.sheet.macroEnabled.12", "upsert": "true"}
    )

def fetch_recent_filings(days_back=7):
    today  = date.today()
    qtr    = (today.month - 1) // 3 + 1
    cutoff = today - timedelta(days=days_back)
    url    = f"https://www.sec.gov/Archives/edgar/full-index/{today.year}/QTR{qtr}/company.idx"

    log.info(f"Fetching EDGAR index: {url}")
    time.sleep(0.5)

    try:
        r = requests.get(url, headers=SEC_HEADERS, timeout=60)
        log.info(f"Status: {r.status_code}, Size: {len(r.content)} bytes")
        if r.status_code != 200:
            log.error(f"Failed: {r.text[:200]}")
            return []
    except Exception as e:
        log.error(f"Request error: {e}")
        return []

    filings = []
    seen = set()
    lines = r.text.split('\n')
    log.info(f"Total lines: {len(lines)}")

    for line in lines:
        if '424B4' not in line:
            continue
        try:
            parts = re.split(r'\s{2,}', line.strip())
            if len(parts) < 4:
                continue
            company   = parts[0].strip()
            form_type = parts[1].strip()
            filed     = parts[3].strip()
            if form_type != '424B4':
                continue
            if not company or company in seen:
                continue
            filed_date = date.fromisoformat(filed[:10])
            if filed_date < cutoff:
                continue
            seen.add(company)
            filings.append({"entity_name": company, "file_date": filed[:10]})
            log.info(f"  Found: {company} ({filed[:10]})")
        except:
            continue

    log.info(f"424B4 filings in last {days_back} days: {len(filings)}")
    return filings

def research_ticker(company_name, filing_date):
    prompt = f"""Research the company "{company_name}" which filed a 424B4 prospectus on {filing_date}.
Return ONLY a JSON object, no other text:
{{
  "ticker": "exchange ticker symbol or empty string",
  "company": "full legal company name",
  "prospectus_date": "YYYY-MM-DD",
  "lockup_days": 180,
  "sec_source": "SEC 424B4 {filing_date}",
  "sponsor": "primary PE/VC sponsor or insider with ownership %",
  "insider_pct": 0.0,
  "ev_sales": "number (compute: Enterprise Value / trailing revenue). Only 'NM' if truly pre-revenue.",
  "ev_ebitda": "number (compute: Enterprise Value / trailing EBITDA). Only 'NM' if EBITDA is negative/zero.",
  "d_score": 0,
  "modifier": 0,
  "early_release": "No",
  "skip": false,
  "skip_reason": ""
}}
Set skip=true if: debt offering, shelf offering, warrant, SPAC, ETF, mutual fund, no lockup.
D-score: 25=active Form4 sellers,22=multiple Form4,20=early release,15=VC/PE upcoming,
12=lockup expired selling,8=PE/VC upcoming,5=corporate parent,0=no mechanism.
Modifier -5 to +5. insider_pct=% held by insiders. Numbers not strings for numeric fields."""

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
        response = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=1000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}]
        )
        text = "".join(b.text for b in response.content if b.type == "text")
        s = text.find("{"); e = text.rfind("}")
        if s == -1: return None
        parsed = json.loads(text[s:e+1])
        log.info(f"  -> ticker={parsed.get('ticker')} skip={parsed.get('skip')}")
        return parsed
    except Exception as e:
        log.error(f"Research error: {e}")
        return None

def log_entry(params, score, tier):
    try:
        supabase.table("lockup_entries").insert({
            "ticker": params["ticker"], "company": params["company"],
            "prospectus_date": params["prospectus_date"], "lockup_days": int(params["lockup_days"]),
            "insider_pct": float(params["insider_pct"]), "ev_sales": str(params["ev_sales"]),
            "ev_ebitda": str(params["ev_ebitda"]), "d_score": float(params["d_score"]),
            "modifier": float(params["modifier"]), "score": score, "tier": tier,
            "sec_source": params.get("sec_source",""), "sponsor": params.get("sponsor",""),
            "early_release": params.get("early_release","No"),
        }).execute()
        log.info(f"Logged {params['ticker']}")
    except Exception as e:
        log.error(f"Supabase log error: {e}")

def get_workbook_tickers(file_bytes):
    """Tickers currently in the workbook (Scoring Model = source of truth)."""
    from openpyxl import load_workbook as _lwb
    from openpyxl.cell.cell import MergedCell as _MC
    wb = _lwb(io.BytesIO(file_bytes), keep_vba=True)
    ws = wb["Scoring Model"]
    out = set()
    for r in range(5, 300):
        c = ws.cell(row=r, column=3)
        if isinstance(c, _MC):
            continue
        if not c.value:
            break
        out.add(str(c.value).upper().strip())
    return out

def reconcile_table_to_workbook(file_bytes):
    """Backfill ONLY explicitly listed names that were logged to lockup_entries
    but never landed in the workbook. The table is append-only history — it also
    holds names deliberately deleted from the workbook (old duplicates, skips),
    so a blanket table-vs-workbook diff would re-add junk. Newest row per ticker
    wins. Returns (file_bytes, n_backfilled)."""
    BACKFILL_TICKERS = {"BLSM", "ATTO", "APMD", "CISS", "AGCC"}
    wb_tickers = get_workbook_tickers(file_bytes)
    todo = BACKFILL_TICKERS - wb_tickers
    if not todo:
        return file_bytes, 0
    res = supabase.table("lockup_entries").select("*") \
        .order("created_at", desc=True).execute()
    latest = {}
    for row in (res.data or []):
        t = str(row.get("ticker", "")).upper().strip()
        if t in todo and t not in latest:
            latest[t] = row
    log.info(f"Reconcile: backfilling {sorted(latest.keys())}")
    from lockup_engine import add_name_to_workbook
    n = 0
    for t, row in latest.items():
        params = {
            "ticker": row["ticker"], "company": row["company"],
            "prospectus_date": row["prospectus_date"], "lockup_days": row["lockup_days"],
            "insider_pct": row["insider_pct"], "ev_sales": row["ev_sales"],
            "ev_ebitda": row["ev_ebitda"], "d_score": row["d_score"],
            "modifier": row["modifier"], "sec_source": row.get("sec_source", ""),
            "sponsor": row.get("sponsor", ""), "early_release": row.get("early_release", "No"),
        }
        try:
            file_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, params)
            n += 1
            log.info(f"  Backfilled {row['ticker']} (score {score}, {tier})")
        except Exception as e:
            log.error(f"  Backfill failed for {row['ticker']}: {e}")
    return file_bytes, n

def sync_historical_backtest(file_bytes):
    """Rebuild Historical Backtest Section 2 from the Short Screen every run:
    one contiguous row per ticker, all formulas self-row referenced. Prevents
    the row-drift / missing-formula corruption from recurring."""
    from openpyxl import load_workbook as _lwb
    from openpyxl.styles import Font, Alignment
    from copy import copy

    wb = _lwb(io.BytesIO(file_bytes), keep_vba=True)
    hb, ss = wb["Historical Backtest"], wb["Short Screen"]

    # kill any stray merges inside the data area (caused the PTRN/KARD corruption)
    for m in [m for m in list(hb.merged_cells.ranges) if m.min_row >= 12]:
        hb.unmerge_cells(str(m))

    # preserve existing static data (sponsor text is richest here)
    existing = {}
    for r in range(12, 200):
        t = hb.cell(row=r, column=2).value
        if not t:
            continue
        t = str(t).strip()
        if t.startswith(("NOTE", "SECTION")):
            continue
        existing[t] = {"company": hb.cell(row=r, column=3).value,
                       "date": hb.cell(row=r, column=4).value,
                       "sponsor": hb.cell(row=r, column=5).value,
                       "float": hb.cell(row=r, column=6).value}

    # sponsor fallback from lockup_entries
    sponsors = {}
    try:
        res = supabase.table("lockup_entries").select("ticker,sponsor").execute()
        sponsors = {str(x["ticker"]).upper(): x.get("sponsor") for x in (res.data or [])}
    except Exception as e:
        log.error(f"Sponsor lookup failed: {e}")

    universe = []
    for r in range(5, 300):
        t = ss.cell(row=r, column=3).value
        if not t:
            break
        universe.append({"ticker": str(t).strip(),
                         "company": ss.cell(row=r, column=4).value,
                         "lockup": ss.cell(row=r, column=5).value,
                         "float": ss.cell(row=r, column=9).value,
                         "thesis": ss.cell(row=r, column=14).value or ""})

    tpl  = {c: copy(hb.cell(row=12, column=c)._style) for c in range(2, 18)}
    tfmt = {c: hb.cell(row=12, column=c).number_format for c in range(2, 18)}

    for r in range(12, 12 + max(len(universe), len(existing)) + 20):
        for c in range(2, 18):
            hb.cell(row=r, column=c).value = None

    r = 12
    for u in universe:
        ex = existing.get(u["ticker"], {})
        sp = (ex.get("sponsor") or sponsors.get(u["ticker"].upper())
              or u["thesis"][:80].strip())
        vals = {2: u["ticker"], 3: ex.get("company") or u["company"],
                4: ex.get("date") or u["lockup"], 5: sp,
                6: ex.get("float") if ex.get("float") is not None else u["float"],
                7: f'=IFERROR(_xlfn.SINGLE(_xlfn.STOCKHISTORY("{u["ticker"]}",N{r},N{r},0,0,1)),"")',
                8: f'=IFERROR(_xlfn.SINGLE(_xlfn.STOCKHISTORY("{u["ticker"]}",O{r},O{r},0,0,1)),"")',
                9: f'=IFERROR(_xlfn.SINGLE(_xlfn.STOCKHISTORY("{u["ticker"]}",P{r},P{r},0,0,1)),"")',
                10: f'=IFERROR(_xlfn.SINGLE(_xlfn.STOCKHISTORY("{u["ticker"]}",Q{r},Q{r},0,0,1)),"")',
                11: f'=IFERROR(H{r}/G{r}-1,"")', 12: f'=IFERROR(I{r}/G{r}-1,"")',
                13: f'=IFERROR(J{r}/G{r}-1,"")',
                14: f'=WORKDAY(D{r},-1)', 15: f'=WORKDAY(D{r},1)',
                16: f'=WORKDAY(D{r},5)', 17: f'=WORKDAY(D{r},10)'}
        for c, v in vals.items():
            cell = hb.cell(row=r, column=c)
            cell.value = v
            cell._style = copy(tpl[c])
            cell.number_format = tfmt[c]
        r += 1

    nc = hb.cell(row=r + 1, column=2)
    nc.value = ("NOTE: STOCKHISTORY() requires Excel 365 with an active market data feed. "
                "Price cells auto-populate once the lockup date passes; % Chg columns compute "
                "from T-1 close. Rows sync automatically from the Short Screen on every scan.")
    nc.font = Font(italic=True, color="FF808080", size=9, name="Calibri")
    nc.alignment = Alignment(horizontal="left")

    out = io.BytesIO()
    wb.save(out)
    log.info(f"Historical Backtest synced: {len(universe)} rows")
    return out.getvalue()

LIQ_GATE_MM = 5.0  # $MM 30-day avg dollar volume; matches Scoring Model header "≥$5MM"

def resolve_liquidity_gate(file_bytes):
    """Force every Scoring Model liquidity gate (col S) to PASS or FAIL — never
    'pending'. If ADV (col R) is blank, retry via yfinance on whatever trading
    days exist; if still no data, FAIL (unverifiable liquidity = untradeable
    for a short screen)."""
    from openpyxl import load_workbook as _lwb
    wb = _lwb(io.BytesIO(file_bytes), keep_vba=True)
    ws = wb["Scoring Model"]
    changed = 0
    for r in range(5, 300):
        t = ws.cell(row=r, column=3).value
        if not t:
            break
        adv = ws.cell(row=r, column=18).value  # R
        if adv is None or str(adv).strip() == "":
            adv = _fetch_adv_mm(str(t).strip())
            if adv is not None:
                ws.cell(row=r, column=18).value = round(adv, 1)
        gate_cell = ws.cell(row=r, column=19)  # S
        try:
            gate = "PASS" if float(adv) >= LIQ_GATE_MM else "FAIL"
        except (TypeError, ValueError):
            gate = "FAIL"  # no data obtainable
        if gate_cell.value != gate:
            gate_cell.value = gate
            changed += 1
            log.info(f"  Liq gate {t}: ADV={adv} -> {gate}")
    out = io.BytesIO()
    wb.save(out)
    log.info(f"Liquidity gate resolved: {changed} cells updated")
    return out.getvalue()

def _fetch_adv_mm(ticker):
    """30-day (or available-days) average daily dollar volume in $MM, or None."""
    try:
        import yfinance as yf
        h = yf.Ticker(ticker).history(period="35d")
        if h is None or h.empty:
            return None
        dv = (h["Close"] * h["Volume"]).dropna()
        if len(dv) == 0:
            return None
        return float(dv.tail(30).mean()) / 1e6
    except Exception as e:
        log.error(f"  ADV fetch failed for {ticker}: {e}")
        return None

# ── Universe rules: exchange + market cap ────────────────────────────────────
ALLOWED_EXCHANGES = {
    "NYQ": "NYSE", "NMS": "Nasdaq", "NGM": "Nasdaq", "NCM": "Nasdaq",
    "TOR": "TSX",
}
MIN_MCAP_USD = 500e6  # $500M floor — names below are benched, not shown

_profile_cache = {}

def _get_profile(ticker):
    """(exchange_code, market_cap_usd) via yfinance; (None, None) if unavailable."""
    t = str(ticker).upper().strip()
    if t in _profile_cache:
        return _profile_cache[t]
    exch, mcap = None, None
    try:
        import yfinance as yf
        info = yf.Ticker(t).info or {}
        exch = info.get("exchange")
        mcap = info.get("marketCap")
        if mcap is None:
            p, so = info.get("regularMarketPrice"), info.get("sharesOutstanding")
            if p and so:
                mcap = p * so
    except Exception as e:
        log.error(f"  Profile fetch failed for {t}: {e}")
    _profile_cache[t] = (exch, mcap)
    return exch, mcap

def qualifies(ticker):
    """(bool, reason). Fails on non-NYSE/Nasdaq/TSX exchange, mcap < $500M,
    or missing data (benched until data appears — new IPOs re-check daily)."""
    exch, mcap = _get_profile(ticker)
    if exch is None and mcap is None:
        return False, "no exchange/mcap data yet"
    if exch not in ALLOWED_EXCHANGES:
        return False, f"exchange {exch} not NYSE/Nasdaq/TSX"
    if mcap is None:
        return False, "no market cap data yet"
    if mcap < MIN_MCAP_USD:
        return False, f"mcap ${mcap/1e6:,.0f}M < $500M"
    return True, f"{ALLOWED_EXCHANGES[exch]} ${mcap/1e9:,.1f}B"

import re as _re

def _rewrite_row_refs(formula, new_row):
    """Rewrite every single-cell row reference in a self-row formula to new_row."""
    return _re.sub(r"([A-Z]{1,2})\d+(?![0-9])", lambda m: f"{m.group(1)}{new_row}", formula)

def enforce_universe_rules(file_bytes):
    """Remove non-qualifying and duplicate tickers from all display sheets,
    park them on hidden _bench, and promote benched names that now qualify
    (re-added from lockup_entries via lockup_engine). Rewrites self-row
    formulas and ranks after row deletion. Returns file_bytes."""
    from openpyxl import load_workbook as _lwb
    wb = _lwb(io.BytesIO(file_bytes), keep_vba=True)
    ss, sm = wb["Short Screen"], wb["Scoring Model"]

    # current universe from Scoring Model
    tickers, seen = [], set()
    dupes = set()
    for r in range(5, 300):
        t = sm.cell(row=r, column=3).value
        if not t:
            break
        t = str(t).upper().strip()
        if t in seen:
            dupes.add(t)
        seen.add(t)
        tickers.append(t)

    verdicts = {t: qualifies(t) for t in set(tickers)}
    drop = {t for t, (ok, _) in verdicts.items() if not ok}
    for t in sorted(set(tickers)):
        ok, reason = verdicts[t]
        log.info(f"  Universe {t}: {'KEEP' if ok else 'BENCH'} ({reason})")
    if dupes:
        log.info(f"  Duplicates to collapse: {sorted(dupes)}")

    # ---- bench sheet
    if "_bench" not in wb.sheetnames:
        b = wb.create_sheet("_bench")
        b.sheet_state = "hidden"
        b["A1"], b["B1"], b["C1"] = "ticker", "benched_at", "reason"
    bench = wb["_bench"]
    benched_now = {str(bench.cell(row=r, column=1).value).upper().strip()
                   for r in range(2, 300) if bench.cell(row=r, column=1).value}

    # ---- delete rows for dropped/duplicate tickers, sheet by sheet
    def _is_ticker(v):
        s = str(v or "").strip()
        return 1 <= len(s) <= 6 and s == s.upper() and " " not in s and s.isascii() and any(ch.isalpha() for ch in s)

    def purge(ws, tick_col, start=5):
        # find last TICKER row (rubric/doc text below the table must not count)
        # so blank spacer rows mid-table don't truncate the scan
        last = 0
        for rr in range(start, ws.max_row + 1):
            if _is_ticker(ws.cell(row=rr, column=tick_col).value):
                last = rr
        removed, kept_seen = 0, set()
        r = start
        while r <= last:
            t = ws.cell(row=r, column=tick_col).value
            if not _is_ticker(t):
                ws.delete_rows(r, 1)   # compact blank spacer
                last -= 1
                removed += 1
                continue
            t = str(t).upper().strip()
            if t in drop or t in kept_seen:
                ws.delete_rows(r, 1)
                last -= 1
                removed += 1
                continue
            kept_seen.add(t)
            r += 1
        return removed

    purge(ss, 3); purge(sm, 3)
    if "Lockup Verification" in wb.sheetnames:
        purge(wb["Lockup Verification"], 2)
    if "Sources" in wb.sheetnames:
        srcs = wb["Sources"]
        r = 5
        while r <= srcs.max_row:
            t = srcs.cell(row=r, column=2).value
            if t and str(t).upper().strip() in drop:
                srcs.delete_rows(r, 1)
                continue
            r += 1

    # ---- rewrite self-row formulas + ranks (openpyxl does not adjust formula
    #      text on delete_rows)
    for r in range(5, 300):
        if not _is_ticker(ss.cell(row=r, column=3).value):
            break
        ss.cell(row=r, column=2).value = r - 4          # rank
        f = ss.cell(row=r, column=6).value               # Days to Lockup
        if isinstance(f, str) and f.startswith("="):
            ss.cell(row=r, column=6).value = f"=E{r}-TODAY()"
    for r in range(5, 300):
        if not _is_ticker(sm.cell(row=r, column=3).value):
            break
        sm.cell(row=r, column=2).value = r - 4
        for c in (7, 8, 11, 14):                         # G,H,K,N
            f = sm.cell(row=r, column=c).value
            if isinstance(f, str) and f.startswith("="):
                sm.cell(row=r, column=c).value = _rewrite_row_refs(f, r)
    if "Lockup Verification" in wb.sheetnames:
        lv = wb["Lockup Verification"]
        for r in range(5, 300):
            if not lv.cell(row=r, column=2).value:
                break
            f = lv.cell(row=r, column=7).value           # Status
            if isinstance(f, str) and f.startswith("="):
                lv.cell(row=r, column=7).value = _rewrite_row_refs(f, r)

    # ---- record newly benched
    next_b = 2
    while bench.cell(row=next_b, column=1).value:
        next_b += 1
    for t in sorted(drop - benched_now):
        bench.cell(row=next_b, column=1).value = t
        bench.cell(row=next_b, column=2).value = date.today().isoformat()
        bench.cell(row=next_b, column=3).value = verdicts[t][1]
        next_b += 1

    out = io.BytesIO()
    wb.save(out)
    file_bytes = out.getvalue()
    log.info(f"Universe enforced: {len(drop)} benched, dupes collapsed: {sorted(dupes) or 'none'}")

    # ---- promotion: benched names that now qualify, with data in lockup_entries
    wb_tickers = get_workbook_tickers(file_bytes)
    promote = [t for t in benched_now
               if t not in wb_tickers and t not in drop and qualifies(t)[0]]
    if promote:
        try:
            res = supabase.table("lockup_entries").select("*")                 .order("created_at", desc=True).execute()
            latest = {}
            for row in (res.data or []):
                t = str(row.get("ticker", "")).upper().strip()
                if t in promote and t not in latest:
                    latest[t] = row
            from lockup_engine import add_name_to_workbook
            for t, row in latest.items():
                params = {k: row[k] for k in ("ticker", "company", "prospectus_date",
                          "lockup_days", "insider_pct", "ev_sales", "ev_ebitda",
                          "d_score", "modifier")}
                params.update(sec_source=row.get("sec_source", ""),
                              sponsor=row.get("sponsor", ""),
                              early_release=row.get("early_release", "No"))
                try:
                    file_bytes, s, ti, _ = add_name_to_workbook(file_bytes, params)
                    log.info(f"  Promoted {t} from bench (score {s})")
                except Exception as e:
                    log.error(f"  Promotion failed for {t}: {e}")
        except Exception as e:
            log.error(f"Promotion lookup failed: {e}")
    return file_bytes


def run():
    log.info("=== Starting lockup scan ===")

    filings = fetch_recent_filings(days_back=7)
    if not filings:
        log.info("No filings found.")
        return

    try:
        file_bytes = download_workbook()
    except Exception as e:
        log.error(f"Workbook download failed: {e}")
        return

    # Pre-filter obvious non-IPO lockups before hitting Claude API
    SKIP_KEYWORDS = [
        'acquisition corp', 'acquisition inc', 'holdings xi', 'holdings x ',
        'holdings ix', 'holdings viii', 'holdings vii', 'holdings vi',
        'ventures acquisition', 'equity partners', 'capital solutions',
        'spac', 'blank check', 'wilco', 'tianci',
    ]

    def is_likely_ipo(company_name):
        name_lower = company_name.lower()
        for kw in SKIP_KEYWORDS:
            if kw in name_lower:
                log.info(f"  Pre-filter skipped: {company_name} (matched: {kw})")
                return False
        return True

    filtered_filings = [f for f in filings if is_likely_ipo(f["entity_name"])]
    log.info(f"After pre-filter: {len(filtered_filings)} of {len(filings)} filings remain")

    MAX_CLAUDE_CALLS = 3  # hard cap to control API cost
    candidates = []
    seen_tickers = set()

    # Robust dedup: seed with tickers ALREADY in the workbook (source of truth),
    # not the lockup_entries table. The table can contain names that never landed
    # in the workbook (failed add/upload) — those must remain eligible so the
    # reconcile step can backfill them.
    try:
        seen_tickers |= get_workbook_tickers(file_bytes)
        log.info(f"Dedup seeded with {len(seen_tickers)} existing workbook tickers")
    except Exception as e:
        log.error(f"Could not seed dedup from workbook: {e}")

    # Backfill any table entries missing from the workbook (uses stored research,
    # zero Claude calls). Fixes orphans like BLSM/ATTO/APMD/CISS/AGCC.
    try:
        file_bytes, backfilled = reconcile_table_to_workbook(file_bytes)
        if backfilled:
            seen_tickers |= get_workbook_tickers(file_bytes)
            upload_workbook(file_bytes)
            log.info(f"Reconcile: {backfilled} names backfilled and uploaded")
    except Exception as e:
        log.error(f"Reconcile failed: {e}")

    claude_calls = 0
    for f in filtered_filings[:10]:
        if claude_calls >= MAX_CLAUDE_CALLS:
            log.info(f"Claude call cap ({MAX_CLAUDE_CALLS}) reached — stopping research")
            break
        log.info(f"Researching: {f['entity_name']}")
        data = research_ticker(f["entity_name"], f["file_date"])
        claude_calls += 1
        if not data: continue
        if data.get("skip"):
            log.info(f"  Skipped: {data.get('skip_reason')}")
            continue
        ticker = data.get("ticker","").upper().strip()
        if not ticker or ticker in ("N/A","TBD",""): continue
        if ticker in seen_tickers:
            log.info(f"  {ticker} duplicate this run, skipping")
            continue
        seen_tickers.add(ticker)
        # NOTE: no lockup_entries check here — the table is NOT the source of
        # truth for what's in the workbook. seen_tickers (workbook-seeded)
        # already handles dedup; a table row without a workbook row is an
        # orphan handled by reconcile_table_to_workbook().
        ok, reason = qualifies(ticker)
        if not ok:
            log.info(f"  {ticker} fails universe rules ({reason}) — benched, not added")
            continue
        insider = float(data.get("insider_pct") or 0)
        score, tier = calc_score(insider, round(100-insider,2),
                                  data.get("ev_sales","NM"), data.get("ev_ebitda","NM"),
                                  float(data.get("d_score") or 0), float(data.get("modifier") or 0))
        data["_score"] = score; data["_tier"] = tier
        candidates.append(data)
        time.sleep(1)

    candidates.sort(key=lambda x: x["_score"], reverse=True)
    candidates = candidates[:MAX_PER_RUN]
    log.info(f"Adding {len(candidates)} candidates")

    added = 0
    for data in candidates:
        try:
            from lockup_engine import add_name_to_workbook
            file_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, data)
            log_entry(data, score, tier)
            added += 1
            log.info(f"Added {data['ticker']}")
        except Exception as e:
            log.error(f"Failed to add {data['ticker']}: {e}")

    if added > 0:
        upload_workbook(file_bytes)

    # Update momentum (price + Form 4) for all tickers — free, no Claude API
    log.info("Updating momentum in thesis cells...")
    try:
        from movement import update_momentum
        latest = download_workbook()
        latest = update_momentum(latest)
        upload_workbook(latest)
        log.info("Momentum update complete.")
    except Exception as e:
        log.error(f"Momentum update failed: {e}")

    # Keep Historical Backtest in lockstep with the Short Screen (self-row
    # formulas, contiguous rows, no merged-cell corruption).
    try:
        latest = download_workbook()
        latest = enforce_universe_rules(latest)
        latest = sync_historical_backtest(latest)
        latest = resolve_liquidity_gate(latest)
        upload_workbook(latest)
    except Exception as e:
        log.error(f"Historical Backtest sync failed: {e}")

    log.info(f"=== Done. {added} names added. ===")
