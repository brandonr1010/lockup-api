"""
Sector scoring — Component F (±25 points)
Zero Claude API cost. Uses Yahoo Finance (ETFs) + NewsAPI (free tier).
Updates daily with the scanner.
"""
import requests, logging, time
from datetime import date, timedelta

log = logging.getLogger(__name__)

# Sector ETF mapping — classify tickers manually or by keyword
SECTOR_ETF = {
    "Technology":       "XLK",
    "Healthcare":       "XLV",
    "Financials":       "XLF",
    "Industrials":      "XLI",
    "Consumer":         "XLY",
    "Energy":           "XLE",
    "Real Estate":      "XLRE",
    "Materials":        "XLB",
    "Utilities":        "XLU",
    "Communications":   "XLC",
    "Biotech":          "XBI",
    "Software":         "IGV",
    "Infrastructure":   "PAVE",
    "Defense":          "ITA",
    "Fintech":          "ARKF",
}

# Manual sector classifications for known tickers
TICKER_SECTOR = {
    "MDLN":  "Healthcare",
    "BOBS":  "Consumer",
    "FPS":   "Energy",
    "GENB":  "Biotech",
    "WLTH":  "Fintech",
    "EQPT":  "Industrials",
    "YSS":   "Defense",
    "CDNL":  "Infrastructure",
    "AGBK":  "Financials",
    "LIFE":  "Technology",
    "KARD":  "Fintech",
    "AVEX":  "Biotech",
    "PTRN":  "Consumer",
    "AKTS":  "Technology",
    "ANDG":  "Industrials",
    "MMED":  "Healthcare",
    "PICS":  "Fintech",
    "SHAZ":  "Technology",
    "MANE":  "Healthcare",
    "FCBM":  "Financials",
    "FISN":  "Industrials",
    "BIAF":  "Biotech",
}

# Sentiment keywords
BEARISH = ["recession","downturn","decline","slump","crash","layoffs","bankruptcy",
           "losses","headwinds","slowdown","miss","disappoints","selloff","plunge",
           "warns","cuts guidance","downgrade","tariff","inflation spike"]
BULLISH = ["surge","rally","boom","record","growth","beat","upgrade","outperform",
           "expansion","strong","momentum","breakout","raises guidance","acquisition",
           "partnership","contract win","AI","tailwind","rate cut","recovery"]

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

def get_news_sentiment(sector, newsapi_key=None):
    """
    Get news sentiment for a sector.
    Returns score -5 to +5.
    Falls back to 0 if no NewsAPI key.
    """
    if not newsapi_key:
        return 0

    keywords = {
        "Technology":     "tech sector stocks",
        "Healthcare":     "healthcare sector stocks",
        "Financials":     "financial sector banks stocks",
        "Industrials":    "industrial sector manufacturing",
        "Consumer":       "consumer discretionary retail stocks",
        "Energy":         "energy sector oil gas stocks",
        "Infrastructure": "infrastructure construction stocks",
        "Biotech":        "biotech stocks FDA",
        "Defense":        "defense aerospace stocks",
        "Fintech":        "fintech payments stocks",
        "Software":       "software SaaS stocks",
        "Real Estate":    "real estate REIT stocks",
        "Communications": "communications media stocks",
    }
    query = keywords.get(sector, f"{sector} stocks")

    try:
        r = requests.get(
            "https://newsapi.org/v2/everything",
            params={
                "q": query,
                "sortBy": "publishedAt",
                "pageSize": 10,
                "language": "en",
                "from": (date.today() - timedelta(days=7)).isoformat(),
                "apiKey": newsapi_key,
            },
            timeout=10
        )
        if r.status_code != 200:
            return 0
        articles = r.json().get("articles", [])
        if not articles:
            return 0

        bull_count = 0
        bear_count = 0
        for article in articles:
            text = ((article.get("title") or "") + " " + (article.get("description") or "")).lower()
            bull_count += sum(1 for w in BULLISH if w in text)
            bear_count += sum(1 for w in BEARISH if w in text)

        net = bull_count - bear_count
        # FLIPPED: bullish news = sector tailwind = anti-short = negative
        # bearish news = sector headwind = pro-short = positive
        if net >= 8:    return -5
        elif net >= 5:  return -3
        elif net >= 2:  return -1
        elif net <= -8: return 5
        elif net <= -5: return 3
        elif net <= -2: return 1
        else:           return 0
    except:
        return 0

def calc_sector_score(ticker, newsapi_key=None):
    """
    Returns F score (±10) for a ticker.
    ETF momentum (±5) + news sentiment (±5).
    """
    sector = TICKER_SECTOR.get(ticker.upper())
    if not sector:
        log.info(f"  No sector mapping for {ticker}, F=0")
        return 0, "Unknown"

    etf = SECTOR_ETF.get(sector)
    etf_pct = get_etf_momentum(etf) if etf else 0.0
    time.sleep(0.3)

    # ETF contribution: ±20 — FLIPPED (sector up = negative for short thesis)
    # ETF up = sector tailwind = anti-short = negative F
    # ETF down = sector headwind = pro-short = positive F
    if etf_pct >= 4.0:    etf_score = -20
    elif etf_pct >= 2.5:  etf_score = -14
    elif etf_pct >= 1.0:  etf_score = -7
    elif etf_pct >= 0:    etf_score = -2
    elif etf_pct >= -1.0: etf_score = 2
    elif etf_pct >= -2.5: etf_score = 7
    elif etf_pct >= -4.0: etf_score = 14
    else:                 etf_score = 20

    news_score = get_news_sentiment(sector, newsapi_key)
    f_score = max(-25, min(25, etf_score + news_score))

    log.info(f"  {ticker} sector={sector} ETF={etf}({etf_pct}%→{etf_score}) news={news_score} F={f_score}")
    return f_score, sector

def get_all_sector_scores(tickers, newsapi_key=None):
    """Returns {ticker: (f_score, sector)} for a list of tickers."""
    results = {}
    # Cache ETF calls — one per sector, not per ticker
    etf_cache = {}
    for ticker in tickers:
        sector = TICKER_SECTOR.get(ticker.upper(), "Unknown")
        etf = SECTOR_ETF.get(sector)
        if etf and etf not in etf_cache:
            etf_cache[etf] = get_etf_momentum(etf)
            time.sleep(0.3)

    news_cache = {}
    for ticker in tickers:
        sector = TICKER_SECTOR.get(ticker.upper(), "Unknown")
        etf = SECTOR_ETF.get(sector)
        etf_pct = etf_cache.get(etf, 0.0)

        # FLIPPED: ETF up = sector tailwind = anti-short = negative F
        # ETF down = sector headwind = pro-short = positive F
        if etf_pct >= 4.0:    etf_score = -20
        elif etf_pct >= 2.5:  etf_score = -14
        elif etf_pct >= 1.0:  etf_score = -7
        elif etf_pct >= 0:    etf_score = -2
        elif etf_pct >= -1.0: etf_score = 2
        elif etf_pct >= -2.5: etf_score = 7
        elif etf_pct >= -4.0: etf_score = 14
        else:                 etf_score = 20

        if sector not in news_cache:
            news_cache[sector] = get_news_sentiment(sector, newsapi_key)
        news_score = news_cache.get(sector, 0)

        f_score = max(-25, min(25, etf_score + news_score))
        results[ticker] = (f_score, sector)
        log.info(f"  {ticker}: sector={sector} ETF={etf_pct}% etf_pts={etf_score} news={news_score} F={f_score}")

    return results
