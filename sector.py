"""
Sector + Stock momentum scoring — Component F (±30 points)
Zero Claude API cost. Uses Yahoo Finance (ETFs + stock prices) + NewsAPI (free tier).

F = ETF sector momentum (±15) + Stock momentum (±15, 60% 1W / 40% 1M)
Capped at ±30. (News disabled — NewsAPI free tier is localhost-only.)

Sign convention (high score = strong short):
  Sector ETF up   = tailwind  = anti-short = NEGATIVE
  Sector ETF down = headwind  = pro-short  = POSITIVE
  Stock up        = short failing = anti-short = NEGATIVE
  Stock down      = short working = pro-short  = POSITIVE
  Bullish news    = anti-short = NEGATIVE
  Bearish news    = pro-short  = POSITIVE
"""
import requests, logging, time
from datetime import date, timedelta

log = logging.getLogger(__name__)

SECTOR_ETF = {
    "Technology":     "XLK",
    "Healthcare":     "XLV",
    "Financials":     "XLF",
    "Industrials":    "XLI",
    "Consumer":       "XLY",
    "Energy":         "XLE",
    "Real Estate":    "XLRE",
    "Materials":      "XLB",
    "Utilities":      "XLU",
    "Communications": "XLC",
    "Biotech":        "XBI",
    "Software":       "IGV",
    "Infrastructure": "PAVE",
    "Defense":        "ITA",
    "Fintech":        "ARKF",
}

TICKER_SECTOR = {
    "MDLN": "Healthcare",
    "BOBS": "Consumer",
    "FPS":  "Energy",
    "GENB": "Biotech",
    "WLTH": "Fintech",
    "EQPT": "Industrials",
    "YSS":  "Defense",
    "CDNL": "Infrastructure",
    "AGBK": "Financials",
    "LIFE": "Technology",
    "KARD": "Fintech",
    "AVEX": "Biotech",
    "PTRN": "Consumer",
    "AKTS": "Technology",
    "ANDG": "Industrials",
    "MMED": "Healthcare",
    "PICS": "Fintech",
    "SHAZ": "Technology",
    "MANE": "Healthcare",
    "FCBM": "Financials",
    "FISN": "Industrials",
    "BIAF": "Biotech",
}

BEARISH = ["recession","downturn","decline","slump","crash","layoffs","bankruptcy",
           "losses","headwinds","slowdown","miss","disappoints","selloff","plunge",
           "warns","cuts guidance","downgrade","tariff","inflation spike"]
BULLISH = ["surge","rally","boom","record","growth","beat","upgrade","outperform",
           "expansion","strong","momentum","breakout","raises guidance","acquisition",
           "partnership","contract win","AI","tailwind","rate cut","recovery"]


# ─── ETF SECTOR MOMENTUM (±13) ───────────────────────────────────
def get_etf_momentum(etf_ticker):
    """Get 1W % change for a sector ETF via Yahoo Finance."""
    try:
        import yfinance as yf
        hist = yf.Ticker(etf_ticker).history(period="10d")
        if hist.empty or len(hist) < 5:
            return 0.0
        cur = hist['Close'].iloc[-1]
        week_ago = hist['Close'].iloc[-6] if len(hist) >= 6 else hist['Close'].iloc[0]
        return round((cur - week_ago) / week_ago * 100, 1)
    except:
        return 0.0


def _etf_to_score(etf_pct):
    """ETF weekly % → ±15. FLIPPED: up = anti-short = negative."""
    if etf_pct >= 4.0:    return -15
    elif etf_pct >= 2.5:  return -11
    elif etf_pct >= 1.0:  return -6
    elif etf_pct >= 0:    return -2
    elif etf_pct >= -1.0: return 2
    elif etf_pct >= -2.5: return 6
    elif etf_pct >= -4.0: return 11
    else:                 return 15


# ─── STOCK MOMENTUM (±13, blended 60% 1W / 40% 1M) ───────────────
def _stock_1w_to_score(w1):
    """Stock 1-week % → ±15. FLIPPED: up = anti-short = negative."""
    if w1 is None: return 0
    if w1 >= 10:    return -15
    elif w1 >= 6:   return -11
    elif w1 >= 3:   return -6
    elif w1 >= 0:   return -2
    elif w1 >= -3:  return 2
    elif w1 >= -6:  return 6
    elif w1 >= -10: return 11
    else:           return 15


def _stock_1m_to_score(m1):
    """Stock 1-month % → ±15. FLIPPED: up = anti-short = negative."""
    if m1 is None: return 0
    if m1 >= 30:    return -15
    elif m1 >= 20:  return -11
    elif m1 >= 10:  return -6
    elif m1 >= 0:   return -2
    elif m1 >= -10: return 2
    elif m1 >= -20: return 6
    elif m1 >= -30: return 11
    else:           return 15


def _stock_to_score(w1, m1):
    """
    Blend 1W and 1M stock momentum: 60% weekly (early signal), 40% monthly (confirmation).
    If only one is available, use it alone. Returns ±13.
    """
    s_w = _stock_1w_to_score(w1)
    s_m = _stock_1m_to_score(m1)
    if w1 is None and m1 is None:
        return 0
    if w1 is None:
        return int(round(s_m))
    if m1 is None:
        return int(round(s_w))
    blended = 0.6 * s_w + 0.4 * s_m
    return int(round(max(-15, min(15, blended))))


# ─── NEWS SENTIMENT (±4) ─────────────────────────────────────────
def get_news_sentiment(sector, newsapi_key=None):
    """DISABLED: NewsAPI free tier is localhost-only, doesn't work from Railway.
    Always returns 0. F now runs on ETF (±15) + stock momentum (±15)."""
    return 0
    # --- disabled below ---
    if not newsapi_key:
        return 0
    keywords = {
        "Technology":"tech sector stocks","Healthcare":"healthcare sector stocks",
        "Financials":"financial sector banks stocks","Industrials":"industrial sector manufacturing",
        "Consumer":"consumer discretionary retail stocks","Energy":"energy sector oil gas stocks",
        "Infrastructure":"infrastructure construction stocks","Biotech":"biotech stocks FDA",
        "Defense":"defense aerospace stocks","Fintech":"fintech payments stocks",
        "Software":"software SaaS stocks","Real Estate":"real estate REIT stocks",
        "Communications":"communications media stocks",
    }
    query = keywords.get(sector, f"{sector} stocks")
    try:
        r = requests.get(
            "https://newsapi.org/v2/everything",
            params={"q": query, "sortBy": "publishedAt", "pageSize": 10, "language": "en",
                    "from": (date.today() - timedelta(days=7)).isoformat(), "apiKey": newsapi_key},
            timeout=10
        )
        if r.status_code != 200:
            return 0
        articles = r.json().get("articles", [])
        if not articles:
            return 0
        bull = bear = 0
        for a in articles:
            text = ((a.get("title") or "") + " " + (a.get("description") or "")).lower()
            bull += sum(1 for w in BULLISH if w in text)
            bear += sum(1 for w in BEARISH if w in text)
        net = bull - bear
        # FLIPPED + capped ±4
        if net >= 8:    return -4
        elif net >= 4:  return -3
        elif net >= 2:  return -1
        elif net <= -8: return 4
        elif net <= -4: return 3
        elif net <= -2: return 1
        else:           return 0
    except:
        return 0


# ─── COMBINED F SCORE ────────────────────────────────────────────
def compute_f(etf_score, stock_score, news_score):
    """Combine the three components, capped at ±30."""
    return int(max(-30, min(30, etf_score + stock_score + news_score)))


def get_sector_context(tickers, newsapi_key=None):
    """
    Returns {sector: (etf_score, news_score)} cached per sector.
    Stock momentum is added later per-ticker (it's stock-specific, not sector-wide).
    """
    context = {}
    etf_cache = {}
    for ticker in tickers:
        sector = TICKER_SECTOR.get(ticker.upper(), "Unknown")
        if sector in context:
            continue
        etf = SECTOR_ETF.get(sector)
        if etf and etf not in etf_cache:
            etf_cache[etf] = get_etf_momentum(etf)
            time.sleep(0.3)
        etf_pct = etf_cache.get(etf, 0.0)
        etf_score = _etf_to_score(etf_pct)
        news_score = get_news_sentiment(sector, newsapi_key)
        context[sector] = (etf_score, news_score, etf_pct)
    return context


def score_ticker(ticker, w1, m1, context):
    """
    Compute F for one ticker given its stock momentum and cached sector context.
    Returns (f_score, sector, breakdown_dict).
    """
    sector = TICKER_SECTOR.get(ticker.upper(), "Unknown")
    etf_score, news_score, etf_pct = context.get(sector, (0, 0, 0.0))
    stock_score = _stock_to_score(w1, m1)
    f = compute_f(etf_score, stock_score, news_score)
    breakdown = {"sector": sector, "etf_pct": etf_pct, "etf_score": etf_score,
                 "stock_score": stock_score, "news_score": news_score, "f": f}
    log.info(f"  {ticker}: sector={sector} ETF={etf_pct}%({etf_score}) "
             f"stock(1W={w1},1M={m1})={stock_score} news={news_score} F={f}")
    return f, sector, breakdown


# ─── LEGACY BATCH (sector + news only, no stock) ─────────────────
def get_all_sector_scores(tickers, newsapi_key=None):
    """
    Legacy: returns {ticker: (f_score, sector)} using sector ETF + news only (no stock).
    Kept for backward compatibility. Stock-aware scoring uses score_ticker().
    """
    context = get_sector_context(tickers, newsapi_key)
    results = {}
    for ticker in tickers:
        sector = TICKER_SECTOR.get(ticker.upper(), "Unknown")
        etf_score, news_score, etf_pct = context.get(sector, (0, 0, 0.0))
        f = compute_f(etf_score, 0, news_score)
        results[ticker] = (f, sector)
    return results


def calc_sector_score(ticker, newsapi_key=None):
    """Single-ticker sector+news score (no stock momentum). Returns (f, sector)."""
    sector = TICKER_SECTOR.get(ticker.upper())
    if not sector:
        return 0, "Unknown"
    etf = SECTOR_ETF.get(sector)
    etf_pct = get_etf_momentum(etf) if etf else 0.0
    time.sleep(0.3)
    etf_score = _etf_to_score(etf_pct)
    news_score = get_news_sentiment(sector, newsapi_key)
    f = compute_f(etf_score, 0, news_score)
    return f, sector
