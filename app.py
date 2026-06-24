import os, json, base64
from flask import Flask, request, jsonify
from flask_cors import CORS
from anthropic import Anthropic
from supabase import create_client
from lockup_engine import add_name_to_workbook, calc_score

app = Flask(__name__)
CORS(app)

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_KEY"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
anthropic = Anthropic(api_key=ANTHROPIC_KEY)

RESEARCH_PROMPT = """You are a financial research assistant for an IPO lockup expiry short-candidate screen.
Research the ticker {ticker} using SEC EDGAR and financial data sources.

Return ONLY a JSON object with exactly these fields — no other text:
{{
  "ticker": "{ticker}",
  "company": "Full legal company name",
  "prospectus_date": "YYYY-MM-DD of most recent 424B4 or IPO prospectus",
  "lockup_days": 180,
  "sec_source": "e.g. SEC 424B4 Jun 17 2026",
  "sponsor": "Primary PE/VC sponsor or insider seller with ownership %",
  "insider_pct": 0.0,
  "ev_sales": "NM or number like 5.2",
  "ev_ebitda": "NM or number like 18.0",
  "d_score": 0,
  "modifier": 0,
  "early_release": "No",
  "notes": "2-3 sentence thesis and any flags for analyst to verify",
  "flags": ["list of items needing analyst verification"]
}}

D-score: 25=active Form 4 sellers, 22=multiple Form 4 sellers, 20=early lockup triggered,
15=VC/PE overhang no confirmed selling, 12=lockup expired selling evidence,
8=PE/VC upcoming no selling, 5=corporate parent, 0=no mechanism.
Modifier: -5 to +5. Negative=anti-short (insider buying, buyback). Positive=extra selling signals.
insider_pct: % held by insiders/pre-IPO holders not yet distributed."""

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

@app.route("/research", methods=["POST"])
def research():
    data = request.json
    ticker = data.get("ticker", "").upper().strip()
    if not ticker:
        return jsonify({"error": "Ticker required"}), 400

    prompt = RESEARCH_PROMPT.format(ticker=ticker)
    try:
        response = anthropic.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}]
        )
        text = "".join(b.text for b in response.content if b.type == "text")
        start = text.find("{"); end = text.rfind("}")
        parsed = json.loads(text[start:end+1])
        return jsonify({"success": True, "data": parsed})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/add", methods=["POST"])
def add():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file_bytes = request.files["file"].read()
    params = json.loads(request.form.get("params", "{}"))

    try:
        updated_bytes, score, tier, days_out = add_name_to_workbook(file_bytes, params)

        # Log to Supabase
        try:
            supabase.table("lockup_entries").insert({
                "ticker": params["ticker"],
                "company": params["company"],
                "prospectus_date": params["prospectus_date"],
                "lockup_days": params["lockup_days"],
                "insider_pct": params["insider_pct"],
                "ev_sales": str(params["ev_sales"]),
                "ev_ebitda": str(params["ev_ebitda"]),
                "d_score": params["d_score"],
                "modifier": params["modifier"],
                "score": score,
                "tier": tier,
                "sec_source": params.get("sec_source", ""),
                "sponsor": params.get("sponsor", ""),
                "early_release": params.get("early_release", "No"),
            }).execute()
        except Exception:
            pass  # Don't fail the whole request if Supabase logging fails

        encoded = base64.b64encode(updated_bytes).decode("utf-8")
        return jsonify({
            "success": True,
            "file_b64": encoded,
            "score": score,
            "tier": tier,
            "days_out": days_out,
            "ticker": params["ticker"]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/history", methods=["GET"])
def history():
    try:
        result = supabase.table("lockup_entries").select("*").order("created_at", desc=True).limit(50).execute()
        return jsonify({"success": True, "data": result.data})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
