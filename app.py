"""
Stock Analysis Web App — Flask backend
Runs the financial model, computes DCF vs current price,
returns rating + Excel download.
"""

import io
import os
import json
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template
from model import (
    fetch_financials, build_assumptions, build_workbook, last_val, cagr
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

# ── DCF calculation (pure Python, mirrors the Excel model) ───────────────────

def compute_dcf(data, assumptions, n_proj=5):
    """
    Compute DCF implied share price in Python so we can display it in the UI.
    Mirrors the logic in _valuation() in model.py.
    """
    scale      = 1_000_000
    stock      = data["stock"]
    price      = stock.get("price", 0) or 0
    shares_out = stock.get("shares_out", 0) or 0
    market_cap = stock.get("market_cap", 0) or (price * shares_out)
    last_cash  = last_val(data["cash"]) or 0
    last_debt  = last_val(data["lt_debt"]) or 0

    rf    = assumptions["rf_rate"]
    beta  = assumptions["beta"]
    erp   = assumptions["erp"]
    coe   = rf + beta * erp
    g     = assumptions["lt_growth"]
    tax   = assumptions["tax_rate"]
    dep   = assumptions["dep_pct"]
    cx    = assumptions["capex_pct"]
    opm   = assumptions["op_margin"]

    base_rev = last_val(data["revenue"]) or 1
    last_sbc = last_val(data.get("stock_comp", {})) or 0

    # Project FCFs year by year (mid-year convention)
    fcfs   = []
    pv_sum = 0.0
    cur_rev = base_rev
    for i in range(n_proj):
        cur_rev = cur_rev * (1 + assumptions["rev_growth"][i])
        ni_i  = cur_rev * opm * (1 - tax)
        da_i  = cur_rev * dep
        cx_i  = cur_rev * cx
        fcf_i = ni_i + da_i - cx_i + last_sbc
        period = i + 0.5
        disc   = 1 / (1 + coe) ** period
        fcfs.append({"year": i + 1, "fcf": fcf_i / scale, "pv": (fcf_i * disc) / scale})
        pv_sum += fcf_i * disc

    # Terminal value (Gordon Growth)
    last_fcf = fcfs[-1]["fcf"] * scale  # back to raw
    fcf_t1   = last_fcf * (1 + g)
    tv       = fcf_t1 / (coe - g) if coe > g else 0
    pv_tv    = tv / (1 + coe) ** n_proj

    # Bridge to equity value
    equity_value = pv_sum + pv_tv + last_cash - last_debt
    shares_mm    = shares_out / scale if shares_out else 1
    dcf_price    = equity_value / scale / shares_mm if shares_mm else 0

    # Upside / downside
    upside = (dcf_price / price - 1) if price and dcf_price else 0

    # Rating thresholds
    if upside >= 0.30:
        rating = "STRONG BUY"
        rating_color = "#00C853"
        rating_desc  = "Model implies ≥30% upside — significantly undervalued"
        stars = 5
    elif upside >= 0.10:
        rating = "BUY"
        rating_color = "#69F0AE"
        rating_desc  = "Model implies 10–30% upside — moderately undervalued"
        stars = 4
    elif upside >= -0.10:
        rating = "HOLD"
        rating_color = "#FFD600"
        rating_desc  = "Model implies within 10% of current price — fairly valued"
        stars = 3
    elif upside >= -0.25:
        rating = "SELL"
        rating_color = "#FF6D00"
        rating_desc  = "Model implies 10–25% downside — moderately overvalued"
        stars = 2
    else:
        rating = "STRONG SELL"
        rating_color = "#D50000"
        rating_desc  = "Model implies >25% downside — significantly overvalued"
        stars = 1

    return {
        "dcf_price":    round(dcf_price, 2),
        "current_price":round(price, 2),
        "upside_pct":   round(upside * 100, 1),
        "rating":       rating,
        "rating_color": rating_color,
        "rating_desc":  rating_desc,
        "stars":        stars,
        "coe_pct":      round(coe * 100, 2),
        "terminal_growth_pct": round(g * 100, 2),
        "pv_fcfs":      round(pv_sum / scale, 1),
        "pv_tv":        round(pv_tv / scale, 1),
        "equity_value": round(equity_value / scale, 1),
        "fcf_schedule": fcfs,
        "n_proj":       n_proj,
    }

def compute_summary(data, assumptions):
    """Pull key metrics for the results card."""
    stock = data["stock"]
    rev   = data["revenue"]
    rev_list = sorted(rev.items())

    # Revenue CAGR
    rev_cagr = cagr(rev, 3)

    # LTM values
    last_rev  = last_val(rev) or 0
    last_ni   = last_val(data["net_income"]) or 0
    last_ocf  = last_val(data["op_cf"]) or 0
    last_capex= abs(last_val(data["capex"]) or 0)
    last_fcf  = last_ocf - last_capex

    shares    = stock.get("shares_out", 0) or 1
    scale     = 1_000_000

    return {
        "name":          data["name"],
        "ticker":        data["ticker"],
        "sector":        data["sector"],
        "industry":      data["industry"],
        "price":         round(stock.get("price", 0) or 0, 2),
        "market_cap_b":  round((stock.get("market_cap", 0) or 0) / 1e9, 2),
        "ev_b":          round((stock.get("ev", 0) or 0) / 1e9, 2),
        "pe_forward":    round(stock.get("pe_forward", 0) or 0, 1),
        "ev_ebitda":     round(stock.get("ev_ebitda", 0) or 0, 1),
        "beta":          round(stock.get("beta", 0) or 0, 2),
        "week52_high":   round(stock.get("week52_high", 0) or 0, 2),
        "week52_low":    round(stock.get("week52_low", 0) or 0, 2),
        "gross_margin":  round((stock.get("gross_margin", 0) or 0) * 100, 1),
        "op_margin":     round((stock.get("op_margin", 0) or 0) * 100, 1),
        "rev_cagr":      round(rev_cagr * 100, 1),
        "last_rev_b":    round(last_rev / 1e9, 2),
        "last_fcf_b":    round(last_fcf / 1e9, 2),
        "target_price":  round(stock.get("target_price", 0) or 0, 2),
        "analyst_rec":   (stock.get("rec") or "N/A").upper(),
        "analyst_count": int(stock.get("analyst_count", 0) or 0),
        "description":   (stock.get("description") or "")[:400],
        "hist_years":    [yr for yr, _ in rev_list],
        "hist_revenue":  [round(v / 1e9, 2) for _, v in rev_list],
        "proj_rev_growth": [round(g * 100, 1) for g in assumptions["rev_growth"]],
    }

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/analyze", methods=["POST"])
def analyze():
    body   = request.get_json() or {}
    ticker = (body.get("ticker") or "").strip().upper()
    n_proj = int(body.get("years", 5))

    if not ticker:
        return jsonify({"error": "Ticker symbol is required"}), 400
    if n_proj not in (3, 4, 5):
        n_proj = 5

    try:
        # 1. Fetch data
        data = fetch_financials(ticker)

        # 2. Build assumptions
        assumptions = build_assumptions(data, n_proj)

        # 3. DCF + rating
        dcf = compute_dcf(data, assumptions, n_proj)

        # 4. Summary card data
        summary = compute_summary(data, assumptions)

        # 5. Build Excel in memory
        wb     = build_workbook(data, n_proj, assumptions)
        buf    = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # 6. Persist temporarily for download
        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname    = f"{ticker}_Model_{ts}.xlsx"
        tmp_dir  = Path("/tmp/stockapp")
        tmp_dir.mkdir(exist_ok=True)
        xlsx_path = tmp_dir / fname
        with open(xlsx_path, "wb") as f:
            f.write(buf.getvalue())

        # Peer data for display
        peers_out = {}
        for pticker, pd in data.get("peers", {}).items():
            peers_out[pticker] = {
                "name":         pd.get("name", pticker),
                "price":        round(pd.get("price", 0) or 0, 2),
                "pe_forward":   round(pd.get("pe_forward", 0) or 0, 1),
                "ev_ebitda":    round(pd.get("ev_ebitda", 0) or 0, 1),
                "gross_margin": round((pd.get("gross_margin", 0) or 0) * 100, 1),
                "op_margin":    round((pd.get("op_margin", 0) or 0) * 100, 1),
                "revenue_growth": round((pd.get("revenue_growth", 0) or 0) * 100, 1),
            }

        return jsonify({
            "ok":       True,
            "ticker":   ticker,
            "filename": fname,
            "summary":  summary,
            "dcf":      dcf,
            "peers":    peers_out,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/api/download/<filename>")
def download(filename):
    # Sanitize
    filename = Path(filename).name
    path = Path("/tmp/stockapp") / filename
    if not path.exists():
        return jsonify({"error": "File not found"}), 404
    return send_file(
        str(path),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True, port=5000)
