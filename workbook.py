import os, io, json, math, time, requests, logging
from datetime import datetime, timedelta, date
from supabase import create_client

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

SUPABASE_URL  = os.environ["SUPABASE_URL"]
SUPABASE_KEY  = os.environ["SUPABASE_SERVICE_KEY"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
BUCKET        = "workbook"
FILE_NAME     = "Lockup_automation2.xlsm"
MAX_PER_RUN   = 10

# SEC requires this exact format: "Name Email" — no org name
HEADERS = {
    "User-Agent": "Brandon Ross brandonr1010@gmail.com",
    "Accept-Encoding": "gzip, deflate",
    "Host": "data.sec.gov"
}

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

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
    C   = excel_round(cS + cE, 1)
    raw = excel_round(A + B + C + d + e, 1)
    score = int(excel_round(max(0, min(100, raw)), 0))
    tier  = "High" if score>=75 else "Medium" if score>=50 else "Low" if score>=25 else "Minimal"
    return score, tier

def download_workbook():
    log.info("Downloading workbook from Supabase...")
    res = supabase.storage.from_(BUCKET).download(FILE_NAME)
    log.info(f"Downloaded {len(res)} bytes")
    return res

def upload_workbook(file_bytes):
    log.info("Uploading workbook to Supabase...")
    supabase.storage.from_(BUCKET).update(
        FILE_NAME, file_bytes,
        {"content-type": "application/vnd.ms-excel.sheet.macroEnabled.12", "upsert": "true"}
    )
    log.info("Upload complete.")

def already_processed(ticker):
    res = supabase.table("lockup_entries").select("ticker").eq("ticker", ticker).execute()
    return len(res.data) > 0

def fetch_recent_filings(days_back=7):
    """Fetch recent 424B4 filings using EDGAR's data.sec.gov submissions API."""
    log.info(f"Fetching recent 424B4 filings (last {days_back} days)...")
    
    cutoff = date.today() - timedelta(days=days_back)
    
    # Use EDGAR company search to find recent 424B4 filers
    # data.sec.gov is the API endpoint designed for programmatic access
    url = "https://data.sec.gov/submissions/recent.json"
    headers = {
        "User-Agent": "Brandon Ross brandonr1010@gmail.com",
        "Accept-Encoding": "gzip, deflate"
    }
    
    try:
        time.sleep(0.5)  # SEC rate limit: 10 req/sec max
        r = requests.get(url, headers=headers, timeout=30)
        log.info(f"data.sec.gov status: {r.status_code}")
        
        if r.status_code == 200:
            data = r.json()
            filings = []
            seen = set()
            
            recent = data.get("filings", {}).get("recent", {})
            forms       = recent.get("form", [])
            dates       = recent.get("filedAt", recent.get("filingDate", []))
            companies   = recent.get("entityName", recent.get("primaryDocument", []))
            accessions  = recent.get("accessionNumber", [])
            
            for i, form in enumerate(forms):
                if form == "424B4":
                    filed = dates[i] if i < len(dates) else ""
                    company = companies[i] if i < len(companies) else ""
                    try:
                        if filed and date.fromisoformat(filed[:10]) >= cutoff:
                            if company and company not in seen:
                                seen.add(company)
                                filings.append({
                                    "entity_name": company,
                                    "file_date": filed[:10],
                                    "accession": accessions[i] if i < len(accessions) else ""
                                })
                                log.info(f"  Found: {company} ({filed[:10]})")
                    except: pass
            
            log.info(f"Found {len(filings)} 424B4 filings")
            if filings:
                return filings
    except Exception as e:
        log.error(f"data.sec.gov error: {e}")
    
    # Fallback: use EDGAR full text search with correct params
    return fetch_via_efts(cutoff)

def fetch_via_efts(cutoff):
    """EDGAR full text search fallback."""
    log.info("Trying EDGAR full text search...")
    headers = {"User-Agent": "Brandon Ross brandonr1010@gmail.com"}
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "q": '"lock-up"',
        "forms": "424B4",
        "dateRange": "custom",
        "startdt": cutoff.isoformat(),
        "enddt": date.today().isoformat(),
        "size": "20"
    }
    try:
        time.sleep(1)
        r = requests.get(url, params=params, headers=headers, timeout=30)
        log.info(f"EFTS status: {r.status_code} url: {r.url}")
        if r.status_code == 200:
            data = r.json()
            hits = data.get("hits", {}).get("hits", [])
            log.info(f"EFTS hits: {len(hits)}")
            filings = []
            seen = set()
            for h in hits:
                src = h.get("_source", {})
                name = src.get("entity_name", "")
                if name and name not in seen:
                    seen.add(name)
                    filings.append({"entity_name": name, "file_date": src.get("file_date", date.today().isoformat())})
                    log.info(f"  EFTS found: {name}")
            return filings
        else:
            log.error(f"EFTS error {r.status_code}: {r.text[:200]}")
    except Exception as e:
        log.error(f"EFTS error: {e}")
    return []

def research_ticker(company_name, filing_date):
    prompt = f"""You are a financial research assistant for an IPO lockup expiry short-candidate screen.
Research the company "{company_name}" which filed a 424B4 prospectus on {filing_date}.

Return ONLY a JSON object with exactly these fields, no other text:
{{
  "ticker": "exchange ticker symbol or empty string if not found",
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

Set skip=true if: not a standard IPO lockup, SPAC, warrant offering, debt offering, shelf offering, or no lockup.
D-score: 25=active Form4 sellers,22=multiple Form4,20=early release,15=VC/PE upcoming no selling,
12=lockup expired selling evidence,8=PE/VC upcoming,5=corporate parent,0=no mechanism.
Modifier -5 to +5. insider_pct=% held by insiders not yet distributed. Numbers not strings for numeric fields."""

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
        if start == -1:
            log.error(f"No JSON for {company_name}")
            return None
        parsed = json.loads(text[start:end+1])
        log.info(f"Researched {company_name}: ticker={parsed.get('ticker')} skip={parsed.get('skip')}")
        return parsed
    except Exception as e:
        log.error(f"Research error for {company_name}: {e}")
        return None

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
        log.info(f"Logged {params['ticker']} to Supabase")
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
        log.error(f"Failed to download workbook: {e}")
        return

    candidates = []
    for f in filings[:20]:
        name = f["entity_name"]
        log.info(f"Researching: {name}")
        data = research_ticker(name, f["file_date"])
        if not data: continue
        if data.get("skip"):
            log.info(f"Skipping {name}: {data.get('skip_reason')}")
            continue
        ticker = data.get("ticker", "").upper().strip()
        if not ticker or ticker in ("N/A", "TBD", ""):
            log.info(f"No valid ticker for {name}")
            continue
        if already_processed(ticker):
            log.info(f"{ticker} already processed")
            continue
        insider   = float(data.get("insider_pct") or 0)
        float_pct = round(100 - insider, 2)
        score, tier = calc_score(insider, float_pct,
                                  data.get("ev_sales","NM"),
                                  data.get("ev_ebitda","NM"),
                                  float(data.get("d_score") or 0),
                                  float(data.get("modifier") or 0))
        data["_score"] = score
        data["_tier"]  = tier
        candidates.append(data)
        time.sleep(2)

    log.info(f"Total candidates: {len(candidates)}")
    candidates.sort(key=lambda x: x["_score"], reverse=True)
    candidates = candidates[:MAX_PER_RUN]

    added = 0
    for data in candidates:
        ticker = data["ticker"]
        log.info(f"Adding {ticker} score={data['_score']}")
        try:
            from lockup_engine import add_name_to_workbook
            file_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, data)
            log_entry(data, score, tier)
            added += 1
            log.info(f"Added {ticker}")
        except Exception as e:
            log.error(f"Failed to add {ticker}: {e}")

    if added > 0:
        upload_workbook(file_bytes)

    log.info(f"=== Done. Added {added} names. ===")
