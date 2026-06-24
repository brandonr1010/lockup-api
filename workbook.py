import os, io, json, math, time, requests, logging
from datetime import datetime, timedelta, date
from supabase import create_client

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

SUPABASE_URL    = os.environ["SUPABASE_URL"]
SUPABASE_KEY    = os.environ["SUPABASE_SERVICE_KEY"]
ANTHROPIC_KEY   = os.environ["ANTHROPIC_API_KEY"]
BUCKET          = "workbook"
FILE_NAME       = "Lockup_automation2.xlsm"
MAX_PER_RUN     = 10
EDGAR_BASE      = "https://efts.sec.gov/LATEST/search-index?q=%22lock-up%22&dateRange=custom&startdt={start}&enddt={end}&forms=424B4"
EDGAR_SEARCH    = "https://efts.sec.gov/LATEST/search-index?q=%22lock-up+agreement%22&forms=424B4&dateRange=custom&startdt={start}&enddt={end}"
HEADERS         = {"User-Agent": "Lucida Capital research@lucida.com", "Accept": "application/json"}

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# ── Score calculation ─────────────────────────────────────────────
def excel_round(val, decimals=0):
    factor = 10 ** decimals
    return math.floor(val * factor + 0.5) / factor

def calc_score(insider, float_pct, ev_sales, ev_ebitda, d, e):
    A = excel_round(min(insider / max(float_pct, 0.01), 5) / 5 * 30, 1)
    B = excel_round(min(insider / 100, 1) * 25, 1)
    try:    cS = min(float(ev_sales) / 5 * 10, 10)
    except: cS = 0
    try:    cE = min(max(float(ev_ebitda) - 5, 0) / 25 * 10, 10)
    except: cE = 0
    C     = excel_round(cS + cE, 1)
    raw   = excel_round(A + B + C + d + e, 1)
    score = int(excel_round(max(0, min(100, raw)), 0))
    tier  = "High" if score>=75 else "Medium" if score>=50 else "Low" if score>=25 else "Minimal"
    return score, tier

# ── Supabase storage ──────────────────────────────────────────────
def download_workbook():
    log.info("Downloading workbook from Supabase...")
    res = supabase.storage.from_(BUCKET).download(FILE_NAME)
    return res

def upload_workbook(file_bytes):
    log.info("Uploading updated workbook to Supabase...")
    supabase.storage.from_(BUCKET).update(
        FILE_NAME, file_bytes,
        {"content-type": "application/vnd.ms-excel.sheet.macroEnabled.12", "upsert": "true"}
    )
    log.info("Upload complete.")

# ── Already processed check ───────────────────────────────────────
def already_processed(ticker):
    res = supabase.table("lockup_entries").select("ticker").eq("ticker", ticker).execute()
    return len(res.data) > 0

# ── EDGAR: fetch recent 424B4 filings ────────────────────────────
def fetch_recent_filings(days_back=1):
    end   = date.today().isoformat()
    start = (date.today() - timedelta(days=days_back)).isoformat()
    url   = f"https://efts.sec.gov/LATEST/search-index?q=%22lock-up%22&forms=424B4&dateRange=custom&startdt={start}&enddt={end}&hits.hits._source=period_of_report,entity_name,file_date,period_of_report"
    
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        data = r.json()
        hits = data.get("hits", {}).get("hits", [])
        filings = []
        seen = set()
        for h in hits:
            src = h.get("_source", {})
            name = src.get("entity_name", "")
            if name and name not in seen:
                seen.add(name)
                filings.append({
                    "entity_name": name,
                    "file_date": src.get("file_date", ""),
                    "accession": h.get("_id", "")
                })
        log.info(f"Found {len(filings)} filings from EDGAR")
        return filings
    except Exception as e:
        log.error(f"EDGAR fetch error: {e}")
        return []

# ── Claude research ───────────────────────────────────────────────
def research_ticker(company_name, filing_date):
    prompt = f"""You are a financial research assistant for an IPO lockup expiry short-candidate screen.
Research the company "{company_name}" which filed a 424B4 prospectus on {filing_date}.

Return ONLY a JSON object with exactly these fields, no other text:
{{
  "ticker": "exchange ticker symbol",
  "company": "full legal company name",
  "prospectus_date": "YYYY-MM-DD",
  "lockup_days": 180,
  "sec_source": "SEC 424B4 {filing_date}",
  "sponsor": "primary PE/VC sponsor or insider with ownership %",
  "insider_pct": 0.0,
  "ev_sales": "NM or number",
  "ev_ebitda": "NM or number",
  "d_score": 0,
  "modifier": 0,
  "early_release": "No",
  "skip": false,
  "skip_reason": ""
}}

If this is not a standard IPO lockup (e.g. it's a SPAC, warrant, debt offering, or has no lockup), set skip=true and explain in skip_reason.

D-score: 25=active Form 4 sellers, 22=multiple Form 4 sellers, 20=early lockup triggered,
15=VC/PE overhang no confirmed selling, 12=lockup expired selling evidence,
8=PE/VC upcoming no selling, 5=corporate parent, 0=no mechanism.
Modifier: -5 to +5. Negative=anti-short. Positive=extra selling signals.
insider_pct: % held by insiders not yet distributed. Return numbers only, not strings, for numeric fields."""

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}]
        )
        text = "".join(b.text for b in response.content if b.type == "text")
        start = text.find("{"); end = text.rfind("}")
        return json.loads(text[start:end+1])
    except Exception as e:
        log.error(f"Claude research error for {company_name}: {e}")
        return None

# ── Workbook write ────────────────────────────────────────────────
def add_to_workbook(file_bytes, params):
    from lockup_engine import add_name_to_workbook
    return add_name_to_workbook(file_bytes, params)

# ── Log to Supabase ───────────────────────────────────────────────
def log_entry(params, score, tier):
    try:
        supabase.table("lockup_entries").insert({
            "ticker":          params["ticker"],
            "company":         params["company"],
            "prospectus_date": params["prospectus_date"],
            "lockup_days":     int(params["lockup_days"]),
            "insider_pct":     float(params["insider_pct"]),
            "ev_sales":        str(params["ev_sales"]),
            "ev_ebitda":       str(params["ev_ebitda"]),
            "d_score":         float(params["d_score"]),
            "modifier":        float(params["modifier"]),
            "score":           score,
            "tier":            tier,
            "sec_source":      params.get("sec_source", ""),
            "sponsor":         params.get("sponsor", ""),
            "early_release":   params.get("early_release", "No"),
        }).execute()
    except Exception as e:
        log.error(f"Supabase log error: {e}")

# ── Main pipeline ─────────────────────────────────────────────────
def run():
    log.info("=== Starting daily lockup scan ===")

    filings = fetch_recent_filings(days_back=1)
    if not filings:
        log.info("No new filings found today.")
        return

    # Download current workbook
    file_bytes = download_workbook()
    added = 0

    # Score and rank candidates first
    candidates = []
    for f in filings:
        if added >= MAX_PER_RUN:
            break
        name = f["entity_name"]
        log.info(f"Researching: {name}")

        data = research_ticker(name, f["file_date"])
        if not data:
            continue
        if data.get("skip"):
            log.info(f"Skipping {name}: {data.get('skip_reason')}")
            continue

        ticker = data.get("ticker", "").upper().strip()
        if not ticker:
            log.info(f"No ticker found for {name}, skipping")
            continue
        if already_processed(ticker):
            log.info(f"{ticker} already in workbook, skipping")
            continue

        insider = float(data.get("insider_pct", 0))
        float_pct = round(100 - insider, 2)
        score, tier = calc_score(insider, float_pct,
                                  data.get("ev_sales", "NM"),
                                  data.get("ev_ebitda", "NM"),
                                  float(data.get("d_score", 0)),
                                  float(data.get("modifier", 0)))
        data["_score"] = score
        data["_tier"]  = tier
        candidates.append(data)
        time.sleep(2)  # rate limit

    # Sort by score, take top MAX_PER_RUN
    candidates.sort(key=lambda x: x["_score"], reverse=True)
    candidates = candidates[:MAX_PER_RUN]

    for data in candidates:
        ticker = data["ticker"]
        score  = data["_score"]
        tier   = data["_tier"]
        log.info(f"Adding {ticker} (score={score} {tier})")
        try:
            file_bytes, score, tier, days_out = add_to_workbook(file_bytes, data)
            log_entry(data, score, tier)
            added += 1
            log.info(f"Added {ticker} successfully")
        except Exception as e:
            log.error(f"Failed to add {ticker}: {e}")

    if added > 0:
        upload_workbook(file_bytes)
        log.info(f"=== Done. Added {added} names. ===")
    else:
        log.info("=== Done. No new names added. ===")

if __name__ == "__main__":
    run()
