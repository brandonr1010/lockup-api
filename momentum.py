"""
Momentum updater — appends live 1W/1M price + Form 4 activity
to the thesis cell (col N) for each ticker on Short Screen.
Zero Anthropic API usage. Free data only: Yahoo Finance + SEC EDGAR.
"""
import io, time, logging, re, requests
from datetime import date, timedelta
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

log = logging.getLogger(__name__)

HEADERS = {"User-Agent": "Brandon Ross brandonr1010@gmail.com"}
MOMENTUM_TAG = "| Momentum update:"  # tag so we can find/replace the appended line

def get_cik(ticker):
    """Look up CIK from SEC EDGAR company tickers JSON."""
    try:
        r = requests.get(
            "https://data.sec.gov/submissions/company_tickers.json",
            headers=HEADERS, timeout=15
        )
        if r.status_code != 200:
            return None
        data = r.json()
        for k, v in data.items():
            if v.get('ticker', '').upper() == ticker.upper():
                return str(v['cik_str']).zfill(10)
        return None
    except:
        return None

def get_form4_activity(cik, days_back=30):
    """
    Check recent Form 4 filings for a CIK.
    Returns a short string like "2 insider sales" or "No Form 4s filed"
    """
    try:
        url = f"https://data.sec.gov/submissions/CIK{cik}.json"
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        data = r.json()
        recent = data.get('filings', {}).get('recent', {})
        forms = recent.get('form', [])
        dates = recent.get('filedAt', recent.get('filingDate', []))
        cutoff = (date.today() - timedelta(days=days_back)).isoformat()
        form4s = [dates[i] for i in range(len(forms)) if forms[i] == '4' and dates[i] >= cutoff]
        if not form4s:
            return f"No Form 4s past {days_back}d"
        return f"{len(form4s)} Form 4 filing{'s' if len(form4s)>1 else ''} past {days_back}d"
    except:
        return None

def get_price_momentum(ticker):
    """Returns (1W_pct, 1M_pct) via Yahoo Finance."""
    try:
        import yfinance as yf
        hist = yf.Ticker(ticker).history(period="35d")
        if hist.empty or len(hist) < 2:
            return None, None
        cur = hist['Close'].iloc[-1]
        w1 = round((cur - hist['Close'].iloc[-6]) / hist['Close'].iloc[-6] * 100, 1) if len(hist) >= 6 else None
        m1 = round((cur - hist['Close'].iloc[-22]) / hist['Close'].iloc[-22] * 100, 1) if len(hist) >= 22 else None
        return w1, m1
    except:
        return None, None

def fmt_pct(val):
    if val is None: return "N/A"
    return f"+{val}%" if val >= 0 else f"{val}%"

def strip_old_momentum(thesis):
    """Remove previously appended momentum line."""
    if not thesis:
        return thesis
    idx = thesis.find(MOMENTUM_TAG)
    if idx != -1:
        return thesis[:idx].rstrip()
    return thesis

def update_momentum(file_bytes):
    """
    For each ticker in Short Screen:
    1. Pull 1W/1M price from Yahoo Finance
    2. Pull recent Form 4 count from EDGAR
    3. Append/replace momentum line in thesis cell (col N)
    Returns updated file bytes.
    """
    wb = load_workbook(io.BytesIO(file_bytes), keep_vba=True)
    wsSS = wb["Short Screen"]

    # Cache CIK lookup — one batch request
    cik_map = {}
    try:
        r = requests.get(
            "https://data.sec.gov/submissions/company_tickers.json",
            headers=HEADERS, timeout=15
        )
        if r.status_code == 200:
            data = r.json()
            for k, v in data.items():
                cik_map[v.get('ticker','').upper()] = str(v['cik_str']).zfill(10)
            log.info(f"CIK map loaded: {len(cik_map)} tickers")
        time.sleep(0.3)
    except Exception as e:
        log.warning(f"CIK batch load failed: {e}")

    updated = 0
    for r in range(5, 200):
        ticker = wsSS.cell(row=r, column=3).value
        if not ticker: break

        log.info(f"  Updating momentum for {ticker}")

        # Price momentum
        w1, m1 = get_price_momentum(ticker)
        time.sleep(0.3)

        # Form 4 activity
        f4_str = None
        cik = cik_map.get(ticker.upper())
        if cik:
            f4_str = get_form4_activity(cik, days_back=30)
            time.sleep(0.35)

        # Build momentum line
        price_str = f"1W {fmt_pct(w1)} / 1M {fmt_pct(m1)}"
        f4_part = f" | {f4_str}" if f4_str else ""
        momentum_line = f"{MOMENTUM_TAG} {price_str}{f4_part}"

        # Update thesis cell
        thesis_cell = wsSS.cell(row=r, column=14)
        if isinstance(thesis_cell, MergedCell):
            continue
        old_thesis = str(thesis_cell.value) if thesis_cell.value else ""
        clean_thesis = strip_old_momentum(old_thesis)
        thesis_cell.value = f"{clean_thesis}  {momentum_line}" if clean_thesis else momentum_line
        updated += 1
        log.info(f"    {ticker}: {momentum_line}")

    log.info(f"Momentum updated for {updated} tickers")
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
