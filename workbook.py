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
HEADERS       = {"User-Agent": "Lucida Capital research@lucida.com"}

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

def fetch_recent_filings(days_back=1):
    """Fetch recent 424B4 filings from EDGAR using the submissions API."""
    end   = date.today()
    start = end - timedelta(days=max(days_back, 7))  # minimum 7 days to ensure results
    
    log.info(f"Fetching 424B4 filings from {start} to {end}")
    
    # Use EDGAR full text search - correct endpoint
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "q": '"lock-up"',
        "forms": "424B4",
        "dateRange": "custom",
        "startdt": start.isoformat(),
        "enddt": end.isoformat(),
        "hits.hits.total.relation": "eq",
        "_source": "entity_name,file_date,period_of_report",
        "size": 20
    }
    
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=30)
        log.info(f"EDGAR response status: {r.status_code}")
        log.info(f"EDGAR URL: {r.url}")
        
        if r.status_code != 200:
            log.error(f"EDGAR error: {r.text[:300]}")
            return fetch_via_rss()
            
        data = r.json()
        hits = data.get("hits", {}).get("hits", [])
        total = data.get("hits", {}).get("total", {}).get("value", 0)
        log.info(f"EDGAR total results: {total}, returned: {len(hits)}")
        
        if not hits:
            log.info("No hits from EDGAR search, trying RSS fallback...")
            return fetch_via_rss()
        
        filings = []
        seen = set()
        for h in hits:
            src = h.get("_source", {})
            name = src.get("entity_name", "")
            if name and name not in seen:
                seen.add(name)
                filings.append({
                    "entity_name": name,
                    "file_date": src.get("file_date", end.isoformat()),
                })
                log.info(f"  Found: {name} ({src.get('file_date', '')})")
        
        return filings
        
    except Exception as e:
        log.error(f"EDGAR fetch error: {e}")
        return fetch_via_rss()

def fetch_via_rss():
    """Fallback: fetch recent 424B4 filings via EDGAR RSS feed."""
    log.info("Trying EDGAR RSS feed for 424B4 filings...")
    url = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcurrent&type=424B4&dateb=&owner=include&count=40&search_text=&output=atom"
    
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        log.info(f"RSS status: {r.status_code}")
        
        if r.status_code != 200:
            log.error(f"RSS failed: {r.text[:200]}")
            return []
        
        # Parse XML
        import xml.etree.ElementTree as ET
        root = ET.fromstring(r.text)
        ns = {"atom": "http://www.w3.org/2005/Atom"}
        
        filings = []
        seen = set()
        cutoff = date.today() - timedelta(days=7)
        
        for entry in root.findall("atom:entry", ns):
            title = entry.find("atom:title", ns)
            updated = entry.find("atom:updated", ns)
            
            if title is None: continue
            title_text = title.text or ""
            
            # Extract company name from title like "424B4 - Company Name (0001234567) (Filer)"
            if " - " in title_text:
                parts = title_text.split(" - ", 1)
                if len(parts) > 1:
                    company_part = parts[1]
                    # Remove CIK and filer type
                    if "(" in company_part:
                        company_name = company_part[:company_part.rfind("(")].strip()
                    else:
                        company_name = company_part.strip()
                    
                    if company_name and company_name not in seen:
                        file_date = date.today().isoformat()
                        if updated is not None:
                            try:
                                file_date = updated.text[:10]
                            except: pass
                        
                        # Only include recent filings
                        try:
                            if date.fromisoformat(file_date) >= cutoff:
                                seen.add(company_name)
                                filings.append({
                                    "entity_name": company_name,
                                    "file_date": file_date,
                                })
                                log.info(f"  RSS found: {company_name} ({file_date})")
                        except: pass
        
        log.info(f"RSS returned {len(filings)} filings")
        return filings
        
    except Exception as e:
        log.error(f"RSS fetch error: {e}")
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

Set skip=true if: not a standard IPO lockup, SPAC, warrant offering, debt offering, or no lockup agreement.
D-score: 25=active Form4 sellers,22=multiple Form4,20=early release triggered,15=VC/PE upcoming no selling,
12=lockup expired selling evidence,8=PE/VC upcoming,5=corporate parent,0=no mechanism.
Modifier -5 to +5. insider_pct=% held by insiders not yet distributed. Use numbers not strings for numeric fields."""

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
            log.error(f"No JSON for {company_name}: {text[:200]}")
            return None
        parsed = json.loads(text[start:end+1])
        log.info(f"Researched {company_name}: ticker={parsed.get('ticker')} skip={parsed.get('skip')} insider={parsed.get('insider_pct')} d={parsed.get('d_score')}")
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
        insider  = float(data.get("insider_pct") or 0)
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

    log.info(f"Total candidates to add: {len(candidates)}")
    candidates.sort(key=lambda x: x["_score"], reverse=True)
    candidates = candidates[:MAX_PER_RUN]

    added = 0
    for data in candidates:
        ticker = data["ticker"]
        log.info(f"Adding {ticker} score={data['_score']} tier={data['_tier']}")
        try:
            from lockup_engine import add_name_to_workbook
            file_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, data)
            log_entry(data, score, tier)
            added += 1
            log.info(f"Added {ticker} successfully")
        except Exception as e:
            log.error(f"Failed to add {ticker}: {e}")

    if added > 0:
        upload_workbook(file_bytes)

    log.info(f"=== Done. Added {added} names. ===")
