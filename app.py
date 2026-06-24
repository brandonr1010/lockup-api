import os, json, io, threading
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from anthropic import Anthropic
from supabase import create_client
from lockup_engine import add_name_to_workbook
from workbook import run as run_scanner, fetch_recent_filings

app = Flask(__name__)
CORS(app)

SUPABASE_URL  = os.environ["SUPABASE_URL"]
SUPABASE_KEY  = os.environ["SUPABASE_SERVICE_KEY"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
CRON_SECRET   = os.environ.get("CRON_SECRET", "lucida-lockup-2026")
BUCKET        = "workbook"
FILE_NAME     = "Lockup_automation2.xlsm"

supabase  = create_client(SUPABASE_URL, SUPABASE_KEY)
anthropic = Anthropic(api_key=ANTHROPIC_KEY)

RESEARCH_PROMPT = """You are a financial research assistant for an IPO lockup expiry short-candidate screen.
Research the ticker {ticker} using SEC EDGAR and financial data sources.
Return ONLY a JSON object, no other text:
{{
  "ticker": "{ticker}",
  "company": "Full legal company name",
  "prospectus_date": "YYYY-MM-DD",
  "lockup_days": 180,
  "sec_source": "e.g. SEC 424B4 Jun 17 2026",
  "sponsor": "Primary PE/VC sponsor or insider seller with ownership %",
  "insider_pct": 0.0,
  "ev_sales": "NM or number",
  "ev_ebitda": "NM or number",
  "d_score": 0,
  "modifier": 0,
  "early_release": "No",
  "notes": "2-3 sentence thesis",
  "flags": ["items needing verification"]
}}
D-score: 25=active Form4 sellers,22=multiple Form4,20=early release,15=VC/PE no selling,
12=expired selling evidence,8=PE/VC upcoming,5=corporate parent,0=no mechanism.
Modifier -5 to +5. insider_pct=% held by insiders not yet distributed."""

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

@app.route("/test-edgar", methods=["GET"])
def test_edgar():
    try:
        filings = fetch_recent_filings(days_back=7)
        return jsonify({"success": True, "count": len(filings), "filings": filings[:10]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/research", methods=["POST"])
def research():
    ticker = request.json.get("ticker", "").upper().strip()
    if not ticker: return jsonify({"error": "Ticker required"}), 400
    try:
        response = anthropic.messages.create(
            model="claude-sonnet-4-6", max_tokens=1000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": RESEARCH_PROMPT.format(ticker=ticker)}]
        )
        text = "".join(b.text for b in response.content if b.type == "text")
        parsed = json.loads(text[text.find("{"):text.rfind("}")+1])
        return jsonify({"success": True, "data": parsed})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/add", methods=["POST"])
def add():
    params = request.json
    if not params: return jsonify({"error": "No params"}), 400
    try:
        file_bytes = supabase.storage.from_(BUCKET).download(FILE_NAME)
        updated_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, params)
        supabase.storage.from_(BUCKET).update(
            FILE_NAME, updated_bytes,
            {"content-type": "application/vnd.ms-excel.sheet.macroEnabled.12", "upsert": "true"}
        )
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
        except: pass
        return jsonify({"success": True, "score": score, "tier": tier,
                        "days_out": days_out, "ticker": params["ticker"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download", methods=["GET"])
def download():
    try:
        file_bytes = supabase.storage.from_(BUCKET).download(FILE_NAME)
        return send_file(io.BytesIO(file_bytes), download_name=FILE_NAME, as_attachment=True,
                         mimetype="application/vnd.ms-excel.sheet.macroEnabled.12")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/scan", methods=["POST"])
def scan():
    if request.headers.get("X-Cron-Secret","") != CRON_SECRET:
        return jsonify({"error": "Unauthorized"}), 401
    thread = threading.Thread(target=run_scanner)
    thread.daemon = True
    thread.start()
    return jsonify({"success": True, "message": "Scan started"})

@app.route("/history", methods=["GET"])
def history():
    try:
        result = supabase.table("lockup_entries").select("*").order("created_at", desc=True).limit(50).execute()
        return jsonify({"success": True, "data": result.data})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

@app.route("/test-raw", methods=["GET"])
def test_raw():
    import requests
    from datetime import date
    headers = {"User-Agent": "Brandon Ross brandonr1010@gmail.com"}
    today = date.today()
    qtr = (today.month - 1) // 3 + 1
    url = f"https://www.sec.gov/Archives/edgar/full-index/{today.year}/QTR{qtr}/company.idx"
    try:
        r = requests.get(url, headers=headers, timeout=30)
        return jsonify({
            "status": r.status_code,
            "headers": dict(r.headers),
            "preview": r.text[:500],
            "url": url
        })
    except Exception as e:
        return jsonify({"error": str(e)})
