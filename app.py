"""
Equilens — Stock Analysis Web App
"""

import io
import logging
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

# Log startup info immediately — visible in Render logs
log.info(f"Python {sys.version}")
log.info(f"CWD: {os.getcwd()}")
log.info(f"Dir contents: {os.listdir('.')}")

from flask import Flask, request, jsonify, send_file, render_template

# ── Tell Flask exactly where templates live regardless of CWD ─────────────────
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
log.info(f"Template dir: {TEMPLATE_DIR} (exists={TEMPLATE_DIR.exists()})")
if TEMPLATE_DIR.exists():
    log.info(f"Template files: {list(TEMPLATE_DIR.iterdir())}")

app = Flask(__name__, template_folder=str(TEMPLATE_DIR))
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024
app.config["JSON_SORT_KEYS"] = False

# ── Import model ──────────────────────────────────────────────────────────────
try:
    from model import (
        fetch_financials, build_assumptions, build_workbook, last_val, cagr
    )
    log.info("model.py imported OK")
except Exception as e:
    log.error(f"model.py import FAILED: {e}")
    traceback.print_exc()
    fetch_financials = build_assumptions = build_workbook = last_val = cagr = None


@app.route("/health")
@app.route("/healthz")
def health():
    model_ok = fetch_financials is not None
    return jsonify({
        "status":       "ok" if model_ok else "model_import_failed",
        "model_loaded": model_ok,
        "cwd":          os.getcwd(),
        "base_dir":     str(BASE_DIR),
        "template_dir": str(TEMPLATE_DIR),
        "template_exists": TEMPLATE_DIR.exists(),
        "timestamp":    datetime.utcnow().isoformat(),
    }), 200 if model_ok else 503


def safe_float(v, default=0.0):
    try:
        r = float(v)
        return r if r == r else default
    except (TypeError, ValueError):
        return default


def compute_dcf(data, assumptions, n_proj=5):
    scale      = 1_000_000
    stock      = data.get("stock", {})
    price      = safe_float(stock.get("price"), 0)
    shares_out = safe_float(stock.get("shares_out"), 0)
    last_cash  = safe_float(last_val(data.get("cash", {})), 0)
    last_debt  = safe_float(last_val(data.get("lt_debt", {})), 0)

    rf   = safe_float(assumptions.get("rf_rate"), 0.045)
    beta = safe_float(assumptions.get("beta"), 1.0)
    erp  = safe_float(assumptions.get("erp"), 0.055)
    coe  = rf + beta * erp
    g    = safe_float(assumptions.get("lt_growth"), 0.03)
    tax  = safe_float(assumptions.get("tax_rate"), 0.21)
    dep  = safe_float(assumptions.get("dep_pct"), 0.04)
    cx   = safe_float(assumptions.get("capex_pct"), 0.05)
    opm  = safe_float(assumptions.get("op_margin"), 0.15)

    if coe <= g:
        coe = g + 0.02

    base_rev   = max(safe_float(last_val(data.get("revenue", {})), 1), 1)
    last_sbc   = safe_float(last_val(data.get("stock_comp", {})), 0)
    rev_growth = assumptions.get("rev_growth", [0.08] * n_proj)

    fcfs = []; pv_sum = 0.0; cur_rev = base_rev
    for i in range(n_proj):
        g_i     = safe_float(rev_growth[i] if i < len(rev_growth) else 0.05, 0.05)
        cur_rev = cur_rev * (1 + g_i)
        fcf_i   = cur_rev * opm * (1-tax) + cur_rev * dep - cur_rev * cx + last_sbc
        disc    = 1.0 / (1 + coe) ** (i + 0.5)
        fcfs.append({"year": i+1, "fcf": round(fcf_i/scale, 1),
                     "pv": round(fcf_i*disc/scale, 1)})
        pv_sum += fcf_i * disc

    last_fcf_raw = fcfs[-1]["fcf"] * scale
    tv    = last_fcf_raw * (1+g) / (coe - g)
    pv_tv = tv / (1+coe)**n_proj

    equity_value = pv_sum + pv_tv + last_cash - last_debt
    shares_mm    = max(shares_out / scale, 0.001)
    dcf_price    = max(0, equity_value / scale / shares_mm)
    if price > 0:
        dcf_price = min(dcf_price, price * 20)
    upside = (dcf_price / price - 1) if price > 0 else 0

    if   upside >= 0.30:  rating,color,desc,stars = "STRONG BUY",  "#1a6b2e","Model implies ≥30% upside — significantly undervalued",5
    elif upside >= 0.10:  rating,color,desc,stars = "BUY",         "#2d9e52","Model implies 10–30% upside — moderately undervalued",4
    elif upside >= -0.10: rating,color,desc,stars = "HOLD",        "#8a6d10","Model implies within ±10% of current price — fairly valued",3
    elif upside >= -0.25: rating,color,desc,stars = "SELL",        "#c45e1a","Model implies 10–25% downside — moderately overvalued",2
    else:                 rating,color,desc,stars = "STRONG SELL", "#a81e1e","Model implies >25% downside — significantly overvalued",1

    return {
        "dcf_price": round(dcf_price,2), "current_price": round(price,2),
        "upside_pct": round(upside*100,1), "rating": rating,
        "rating_color": color, "rating_desc": desc, "stars": stars,
        "coe_pct": round(coe*100,2), "terminal_growth_pct": round(g*100,2),
        "pv_fcfs": round(pv_sum/scale,1), "pv_tv": round(pv_tv/scale,1),
        "equity_value": round(equity_value/scale,1),
        "fcf_schedule": fcfs, "n_proj": n_proj,
    }


def compute_summary(data, assumptions):
    stock    = data.get("stock", {})
    rev      = data.get("revenue", {})
    rev_list = sorted(rev.items()) if rev else []
    rev_cagr = safe_float(cagr(rev, 3), 0) if rev and cagr else 0
    last_rev  = safe_float(last_val(rev), 0) if last_val else 0
    last_ocf  = safe_float(last_val(data.get("op_cf", {})), 0) if last_val else 0
    last_capex= abs(safe_float(last_val(data.get("capex", {})), 0)) if last_val else 0
    return {
        "name":          data.get("name", "Unknown"),
        "ticker":        data.get("ticker", ""),
        "sector":        data.get("sector", "Unknown"),
        "industry":      data.get("industry", "Unknown"),
        "price":         round(safe_float(stock.get("price")), 2),
        "market_cap_b":  round(safe_float(stock.get("market_cap")) / 1e9, 2),
        "ev_b":          round(safe_float(stock.get("ev")) / 1e9, 2),
        "pe_forward":    round(safe_float(stock.get("pe_forward")), 1),
        "ev_ebitda":     round(safe_float(stock.get("ev_ebitda")), 1),
        "beta":          round(safe_float(stock.get("beta"), 1.0), 2),
        "week52_high":   round(safe_float(stock.get("week52_high")), 2),
        "week52_low":    round(safe_float(stock.get("week52_low")), 2),
        "gross_margin":  round(safe_float(stock.get("gross_margin")) * 100, 1),
        "op_margin":     round(safe_float(stock.get("op_margin")) * 100, 1),
        "rev_cagr":      round(rev_cagr * 100, 1),
        "last_rev_b":    round(last_rev / 1e9, 2),
        "last_fcf_b":    round((last_ocf - last_capex) / 1e9, 2),
        "target_price":  round(safe_float(stock.get("target_price")), 2),
        "analyst_rec":   (stock.get("rec") or "N/A").upper(),
        "analyst_count": int(safe_float(stock.get("analyst_count"))),
        "description":   str(stock.get("description") or "")[:400],
        "hist_years":    [yr for yr, _ in rev_list],
        "hist_revenue":  [round(safe_float(v) / 1e9, 2) for _, v in rev_list],
        "proj_rev_growth": [round(safe_float(g_val) * 100, 1)
                            for g_val in assumptions.get("rev_growth", [])],
    }


@app.route("/")
def index():
    log.info(f"Serving index.html from {TEMPLATE_DIR}")
    return render_template("index.html")


@app.route("/api/analyze", methods=["POST"])
def analyze():
    if fetch_financials is None:
        return jsonify({"error": "Model failed to load. Check server logs."}), 503

    body   = request.get_json(silent=True) or {}
    ticker = str(body.get("ticker") or "").strip().upper()
    n_proj = int(body.get("years") or 5)

    if not ticker:
        return jsonify({"error": "Ticker symbol is required."}), 400
    if len(ticker) > 10:
        return jsonify({"error": f"'{ticker}' is not a valid ticker."}), 400
    if n_proj not in (3, 4, 5):
        n_proj = 5

    log.info(f"[{ticker}] Analysis started — {n_proj}yr")

    try:
        log.info(f"[{ticker}] Fetching financials...")
        data = fetch_financials(ticker)
        log.info(f"[{ticker}] Revenue years: {sorted(data.get('revenue',{}).keys())}")

        log.info(f"[{ticker}] Building assumptions...")
        assumptions = build_assumptions(data, n_proj)

        log.info(f"[{ticker}] Computing DCF...")
        dcf = compute_dcf(data, assumptions, n_proj)
        log.info(f"[{ticker}] DCF=${dcf['dcf_price']} mkt=${dcf['current_price']} → {dcf['rating']}")

        summary = compute_summary(data, assumptions)

        log.info(f"[{ticker}] Building Excel...")
        wb  = build_workbook(data, n_proj, assumptions)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname   = f"{ticker}_Model_{ts}.xlsx"
        tmp_dir = Path("/tmp/stockapp")
        tmp_dir.mkdir(parents=True, exist_ok=True)
        (tmp_dir / fname).write_bytes(buf.getvalue())
        log.info(f"[{ticker}] Excel saved: {fname}")

        peers_out = {}
        for pt, pd in (data.get("peers") or {}).items():
            try:
                peers_out[pt] = {
                    "name":           str(pd.get("name") or pt),
                    "price":          round(safe_float(pd.get("price")), 2),
                    "pe_forward":     round(safe_float(pd.get("pe_forward")), 1),
                    "ev_ebitda":      round(safe_float(pd.get("ev_ebitda")), 1),
                    "gross_margin":   round(safe_float(pd.get("gross_margin")) * 100, 1),
                    "op_margin":      round(safe_float(pd.get("op_margin")) * 100, 1),
                    "revenue_growth": round(safe_float(pd.get("revenue_growth")) * 100, 1),
                }
            except Exception:
                pass

        return jsonify({"ok": True, "ticker": ticker, "filename": fname,
                        "summary": summary, "dcf": dcf, "peers": peers_out})

    except Exception as e:
        log.error(f"[{ticker}] FAILED: {type(e).__name__}: {e}")
        log.error(traceback.format_exc())
        msg = str(e)
        if "yfinance" in msg.lower() or "No module" in msg:
            msg = "yfinance not installed — check requirements.txt."
        elif "CIK not found" in msg or "not found" in msg.lower():
            msg = f"Ticker '{ticker}' not found. Please verify the symbol."
        elif "Connection" in msg or "timeout" in msg.lower():
            msg = "Could not reach Yahoo Finance. Please try again."
        elif "NoneType" in msg:
            msg = f"Insufficient data for '{ticker}'. Try a major US-listed stock."
        return jsonify({"error": msg}), 500


@app.route("/api/download/<filename>")
def download(filename):
    safe_name = Path(filename).name
    if not safe_name.endswith(".xlsx"):
        return jsonify({"error": "Invalid file"}), 400
    path = Path("/tmp/stockapp") / safe_name
    if not path.exists():
        return jsonify({"error": "File not found. Please re-run the analysis."}), 404
    return send_file(str(path), as_attachment=True, download_name=safe_name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "Not found"}), 404

@app.errorhandler(500)
def server_error(e):
    log.error(f"Unhandled 500: {e}")
    log.error(traceback.format_exc())
    return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    log.info(f"Starting Equilens on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
