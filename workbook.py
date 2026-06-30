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

def already_processed(ticker):
    res = supabase.table("lockup_entries").select("ticker").eq("ticker", ticker).execute()
    return len(res.data) > 0

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
        if already_processed(ticker):
            log.info(f"  {ticker} already in workbook")
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

    log.info(f"=== Done. {added} names added. ===")
