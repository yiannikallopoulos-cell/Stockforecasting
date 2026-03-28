"""
Sharp Money — Stock Analysis Web App
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

log.info(f"Python {sys.version}")
log.info(f"CWD: {os.getcwd()}")
log.info(f"Dir: {os.listdir('.')}")

from flask import Flask, request, jsonify, send_file, Response

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024
app.config["JSON_SORT_KEYS"]     = False

try:
    from model import (
        fetch_financials, build_assumptions, build_workbook, last_val, cagr
    )
    log.info("model.py imported OK")
except Exception as e:
    log.error(f"model.py FAILED: {e}")
    traceback.print_exc()
    fetch_financials = build_assumptions = build_workbook = last_val = cagr = None


# ── Diagnostic routes ─────────────────────────────────────────────────────────

@app.route("/health")
def health():
    model_ok = fetch_financials is not None
    return jsonify({
        "status":           "ok" if model_ok else "model_import_failed",
        "model_loaded":     model_ok,
        "python":           sys.version,
        "cwd":              os.getcwd(),
        "template_exists":  TEMPLATE_DIR.exists(),
        "timestamp":        datetime.utcnow().isoformat(),
    }), 200 if model_ok else 503


@app.route("/test")
def test():
    """
    Full diagnostic — checks every dependency and network access.
    Visit /test in browser to see exactly what's working/failing.
    """
    results = {}

    # 1. Check imports
    for pkg in ["yfinance","requests","openpyxl","flask"]:
        try:
            mod = __import__(pkg)
            results[f"import_{pkg}"] = getattr(mod, "__version__", "ok")
        except Exception as e:
            results[f"import_{pkg}"] = f"MISSING: {e}"

    # 2. Check network — Yahoo Finance
    try:
        import requests as req
        r = req.get("https://finance.yahoo.com", timeout=10)
        results["network_yahoo"] = f"OK status={r.status_code}"
    except Exception as e:
        results["network_yahoo"] = f"FAILED: {e}"

    # 3. Check network — SEC EDGAR
    try:
        import requests as req
        r = req.get("https://www.sec.gov/files/company_tickers.json",
                    headers={"User-Agent":"test@test.com"}, timeout=10)
        results["network_sec"] = f"OK status={r.status_code}"
    except Exception as e:
        results["network_sec"] = f"FAILED: {e}"

    # 4. Check yfinance — yfinance 1.2+ manages its own session, do NOT pass one
    try:
        import yfinance as yf
        t    = yf.Ticker("AAPL")
        info = {}
        for attempt in range(3):
            try:
                raw = t.info
                if raw and len(raw) > 5: info = raw; break
            except Exception as e:
                results[f"yfinance_attempt_{attempt+1}"] = str(e)
            __import__("time").sleep(2)
        if not info:
            fi = t.fast_info
            px = getattr(fi,"last_price",None) or getattr(fi,"previous_close",None)
            results["yfinance_AAPL"] = f"OK via fast_info price={px}"
        else:
            px = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
            results["yfinance_AAPL"] = f"OK price={px}"
    except Exception as e:
        results["yfinance_AAPL"] = f"FAILED: {e}"

    # 5. Check model import
    results["model_imported"] = fetch_financials is not None

    # 6. Check /tmp writable
    try:
        p = Path("/tmp/testfile.txt")
        p.write_text("ok")
        p.unlink()
        results["tmp_writable"] = "OK"
    except Exception as e:
        results["tmp_writable"] = f"FAILED: {e}"

    return jsonify(results)


# ── Helpers ───────────────────────────────────────────────────────────────────

def safe_float(v, default=0.0):
    try:
        r = float(v)
        return r if r == r else default
    except (TypeError, ValueError):
        return default


def _run_dcf_scenario(data, assumptions, n_proj, scenario_key="base"):
    """
    Run DCF for a given scenario using per-year projected margins,
    NWC drag, and split CapEx. Returns equity value and FCF schedule.
    """
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

    if coe <= g: coe = g + 0.02

    # Pull scenario-specific projections
    scenarios   = assumptions.get("scenarios", {})
    scen        = scenarios.get(scenario_key, scenarios.get("base", {}))
    rev_growth  = scen.get("rev_growth",  assumptions.get("rev_growth",  [0.08]*n_proj))
    om_proj     = scen.get("op_margin",   assumptions.get("op_margin_proj", [assumptions.get("op_margin", 0.15)]*n_proj))
    capex_proj  = scen.get("capex_pct",   assumptions.get("capex_proj",  [assumptions.get("capex_pct", 0.05)]*n_proj))
    nwc_proj    = assumptions.get("nwc_change_proj", [0.0]*n_proj)

    base_rev  = max(safe_float(last_val(data.get("revenue", {})), 1), 1)
    last_sbc  = safe_float(last_val(data.get("stock_comp", {})), 0)

    fcfs = []; pv_sum = 0.0; cur_rev = base_rev
    for i in range(n_proj):
        g_i   = safe_float(rev_growth[i] if i < len(rev_growth) else 0.05, 0.05)
        opm_i = safe_float(om_proj[i]    if i < len(om_proj)    else 0.15, 0.15)
        cx_i  = safe_float(capex_proj[i] if i < len(capex_proj) else 0.05, 0.05)
        nwc_i = safe_float(nwc_proj[i]   if i < len(nwc_proj)   else 0.0,  0.0)

        cur_rev = cur_rev * (1 + g_i)
        ni_i    = cur_rev * opm_i * (1 - tax)
        da_i    = cur_rev * dep
        cx_val  = cur_rev * cx_i
        nwc_drag= cur_rev * nwc_i          # cash consumed by working capital build
        sbc_i   = last_sbc
        fcf_i   = ni_i + da_i - cx_val - nwc_drag + sbc_i

        disc = 1.0 / (1 + coe) ** (i + 0.5)
        fcfs.append({
            "year":      i + 1,
            "revenue":   round(cur_rev / scale, 1),
            "op_margin": round(opm_i * 100, 1),
            "fcf":       round(fcf_i / scale, 1),
            "pv":        round(fcf_i * disc / scale, 1),
        })
        pv_sum += fcf_i * disc

    # Terminal value using last year FCF (Gordon Growth)
    last_fcf_raw = fcfs[-1]["fcf"] * scale
    tv    = last_fcf_raw * (1 + g) / (coe - g)
    pv_tv = tv / (1 + coe) ** n_proj

    eq_v  = pv_sum + pv_tv + last_cash - last_debt
    shrs  = max(shares_out / scale, 0.001)
    dcf_p = max(0, eq_v / scale / shrs)
    if price > 0:
        dcf_p = min(dcf_p, price * 25)

    return {
        "dcf_price":    round(dcf_p, 2),
        "pv_fcfs":      round(pv_sum / scale, 1),
        "pv_tv":        round(pv_tv / scale, 1),
        "equity_value": round(eq_v / scale, 1),
        "fcf_schedule": fcfs,
        "coe":          coe,
    }


def _rating_from_upside(upside):
    if   upside >= 0.30:  return "STRONG BUY",  "#1a6b2e", "Model implies ≥30% upside — significantly undervalued", 5
    elif upside >= 0.10:  return "BUY",          "#2d9e52", "Model implies 10–30% upside — moderately undervalued", 4
    elif upside >= -0.10: return "HOLD",         "#8a6d10", "Model implies within ±10% of current price — fairly valued", 3
    elif upside >= -0.25: return "SELL",         "#c45e1a", "Model implies 10–25% downside — moderately overvalued", 2
    else:                 return "STRONG SELL",  "#a81e1e", "Model implies >25% downside — significantly overvalued", 1


def compute_dcf(data, assumptions, n_proj=5):
    """
    Run DCF across all three scenarios (bull/base/bear).
    Base case drives the primary rating; scenarios provide the price range.
    """
    stock  = data.get("stock", {})
    price  = safe_float(stock.get("price"), 0)

    # Run all three scenarios
    base_res = _run_dcf_scenario(data, assumptions, n_proj, "base")
    bull_res = _run_dcf_scenario(data, assumptions, n_proj, "bull")
    bear_res = _run_dcf_scenario(data, assumptions, n_proj, "bear")

    dcf_p  = base_res["dcf_price"]
    coe    = base_res["coe"]
    g      = safe_float(assumptions.get("lt_growth"), 0.03)
    upside = (dcf_p / price - 1) if price > 0 else 0

    r, c, d, s = _rating_from_upside(upside)

    # Build sensitivity table: CoE (rows) vs terminal growth (cols)
    coe_range = [coe - 0.02, coe - 0.01, coe, coe + 0.01, coe + 0.02]
    g_range   = [g - 0.01, g, g + 0.01]
    sensitivity = []
    for coe_s in coe_range:
        row_s = []
        for g_s in g_range:
            try:
                tmp_assumptions = dict(assumptions)
                tmp_assumptions["lt_growth"] = g_s
                # Quick terminal value recalc using base FCF
                last_fcf = base_res["fcf_schedule"][-1]["fcf"] * 1_000_000
                tv_s     = last_fcf * (1 + g_s) / (max(coe_s, g_s + 0.001) - g_s)
                pv_tv_s  = tv_s / (1 + coe_s) ** n_proj
                scale    = 1_000_000
                shares   = max(safe_float(stock.get("shares_out"), 1) / scale, 0.001)
                last_cash= safe_float(last_val(data.get("cash", {})), 0) / scale
                last_debt= safe_float(last_val(data.get("lt_debt", {})), 0) / scale
                eq_s     = base_res["pv_fcfs"] + pv_tv_s / scale + last_cash - last_debt
                px_s     = max(0, eq_s / shares)
                row_s.append(round(px_s, 2))
            except Exception:
                row_s.append(None)
        sensitivity.append(row_s)

    return {
        "dcf_price":           dcf_p,
        "current_price":       round(price, 2),
        "upside_pct":          round(upside * 100, 1),
        "rating":              r,
        "rating_color":        c,
        "rating_desc":         d,
        "stars":               s,
        "coe_pct":             round(coe * 100, 2),
        "terminal_growth_pct": round(g * 100, 2),
        "pv_fcfs":             base_res["pv_fcfs"],
        "pv_tv":               base_res["pv_tv"],
        "equity_value":        base_res["equity_value"],
        "fcf_schedule":        base_res["fcf_schedule"],
        "n_proj":              n_proj,
        # Scenario prices
        "bull_price":          bull_res["dcf_price"],
        "bear_price":          bear_res["dcf_price"],
        "bull_upside":         round((bull_res["dcf_price"]/price-1)*100, 1) if price else 0,
        "bear_upside":         round((bear_res["dcf_price"]/price-1)*100, 1) if price else 0,
        # Model metadata
        "stage":               assumptions.get("stage", "unknown"),
        "dol":                 assumptions.get("dol", 1.0),
        "gm_trend":            assumptions.get("gm_trend", 0),
        "om_trend":            assumptions.get("om_trend", 0),
        # Sensitivity table
        "sensitivity_prices":  sensitivity,
        "sensitivity_coe":     [round(c*100,1) for c in coe_range],
        "sensitivity_g":       [round(g_val*100,1) for g_val in g_range],
    }


def compute_summary(data, assumptions):
    stock    = data.get("stock", {})
    rev      = data.get("revenue", {})
    rev_list = sorted(rev.items()) if rev else []
    rev_cagr = safe_float(cagr(rev,3),0) if (rev and cagr) else 0
    lv       = last_val if last_val else (lambda x,d=0: d)
    last_rev  = safe_float(lv(rev), 0)
    last_ocf  = safe_float(lv(data.get("op_cf",{})), 0)
    last_capex= abs(safe_float(lv(data.get("capex",{})), 0))
    return {
        "name":data.get("name","Unknown"),"ticker":data.get("ticker",""),
        "sector":data.get("sector","Unknown"),"industry":data.get("industry","Unknown"),
        "price":round(safe_float(stock.get("price")),2),
        "market_cap_b":round(safe_float(stock.get("market_cap"))/1e9,2),
        "ev_b":round(safe_float(stock.get("ev"))/1e9,2),
        "pe_forward":round(safe_float(stock.get("pe_forward")),1),
        "ev_ebitda":round(safe_float(stock.get("ev_ebitda")),1),
        "beta":round(safe_float(stock.get("beta"),1.0),2),
        "week52_high":round(safe_float(stock.get("week52_high")),2),
        "week52_low":round(safe_float(stock.get("week52_low")),2),
        "gross_margin":round(safe_float(stock.get("gross_margin"))*100,1),
        "op_margin":round(safe_float(stock.get("op_margin"))*100,1),
        "rev_cagr":round(rev_cagr*100,1),
        "last_rev_b":round(last_rev/1e9,2),
        "last_fcf_b":round((last_ocf-last_capex)/1e9,2),
        "target_price":round(safe_float(stock.get("target_price")),2),
        "analyst_rec":(stock.get("rec") or "N/A").upper(),
        "analyst_count":int(safe_float(stock.get("analyst_count"))),
        "description":str(stock.get("description") or "")[:400],
        "hist_years":[yr for yr,_ in rev_list],
        "hist_revenue":[round(safe_float(v)/1e9,2) for _,v in rev_list],
        "proj_rev_growth":[round(safe_float(gv)*100,1) for gv in assumptions.get("rev_growth",[])],
        # ── Quarterly signals ──
        "quarterly_revenue":{k:round(safe_float(v)/1e9,2) for k,v in (data.get("quarterly",{}).get("quarterly_revenue",{}) or {}).items()},
        "quarterly_gross_margin":{k:round(safe_float(v)*100,1) for k,v in (data.get("quarterly",{}).get("quarterly_gross_margin",{}) or {}).items()},
        "quarterly_op_margin":{k:round(safe_float(v)*100,1) for k,v in (data.get("quarterly",{}).get("quarterly_op_margin",{}) or {}).items()},
        "recent_rev_accel":round(safe_float(data.get("quarterly",{}).get("recent_revenue_accel")),2),
        # ── Insider activity ──
        "insider_signal":      data.get("insider",{}).get("insider_signal","neutral"),
        "insider_signal_desc": data.get("insider",{}).get("insider_signal_desc",""),
        "insider_buy_count":   int(data.get("insider",{}).get("insider_buy_count",0) or 0),
        "insider_sell_count":  int(data.get("insider",{}).get("insider_sell_count",0) or 0),
        "insider_transactions":data.get("insider",{}).get("insider_transactions",[])[:5],
        # ── Short interest ──
        "short_pct_float":  safe_float(data.get("short_interest",{}).get("short_pct_float")),
        "short_ratio":      safe_float(data.get("short_interest",{}).get("short_ratio")),
        "short_signal":     data.get("short_interest",{}).get("short_signal","neutral"),
        "short_signal_desc":data.get("short_interest",{}).get("short_signal_desc",""),
        # ── Options IV ──
        "iv_30d":         safe_float(data.get("options_iv",{}).get("iv_30d")),
        "put_call_ratio": safe_float(data.get("options_iv",{}).get("put_call_ratio")),
        "iv_signal":      data.get("options_iv",{}).get("iv_signal","neutral"),
        "iv_signal_desc": data.get("options_iv",{}).get("iv_signal_desc",""),
        # ── Segments ──
        "segment_note":  data.get("segments",{}).get("segment_note",""),
        "has_segments":  data.get("segments",{}).get("has_segments",False),
        "segment_names": list((data.get("segments",{}).get("segments",{}) or {}).keys()),
    }


# ── Routes ────────────────────────────────────────────────────────────────────

# HTML embedded directly — no templates folder needed
INDEX_HTML = open(__file__.replace("app.py","index.html"), encoding="utf-8").read()

@app.route("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


@app.route("/api/analyze", methods=["POST"])
def analyze():
    if fetch_financials is None:
        return jsonify({"error":"Model failed to load. Visit /test for diagnostics."}),503

    body   = request.get_json(silent=True) or {}
    ticker = str(body.get("ticker") or "").strip().upper()
    n_proj = int(body.get("years") or 5)

    if not ticker:
        return jsonify({"error":"Ticker symbol is required."}), 400
    if len(ticker) > 10:
        return jsonify({"error":f"'{ticker}' is not a valid ticker."}), 400
    if n_proj not in (3,4,5):
        n_proj = 5

    log.info(f"[{ticker}] Analysis started {n_proj}yr")

    try:
        log.info(f"[{ticker}] fetch_financials...")
        data = fetch_financials(ticker)
        log.info(f"[{ticker}] revenue={sorted(data.get('revenue',{}).keys())}")

        log.info(f"[{ticker}] build_assumptions...")
        assumptions = build_assumptions(data, n_proj)

        log.info(f"[{ticker}] compute_dcf...")
        dcf = compute_dcf(data, assumptions, n_proj)
        log.info(f"[{ticker}] DCF=${dcf['dcf_price']} mkt=${dcf['current_price']} {dcf['rating']}")

        summary = compute_summary(data, assumptions)

        log.info(f"[{ticker}] build_workbook...")
        wb  = build_workbook(data, n_proj, assumptions)
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)

        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname   = f"{ticker}_Model_{ts}.xlsx"
        tmp_dir = Path("/tmp/stockapp")
        tmp_dir.mkdir(parents=True, exist_ok=True)
        (tmp_dir/fname).write_bytes(buf.getvalue())
        log.info(f"[{ticker}] saved {fname}")

        peers_out = {}
        for pt,pd in (data.get("peers") or {}).items():
            try:
                peers_out[pt] = {
                    "name":           str(pd.get("name") or pt),
                    "price":          round(safe_float(pd.get("price")),2),
                    "pe_forward":     round(safe_float(pd.get("pe_forward")),1),
                    "ev_ebitda":      round(safe_float(pd.get("ev_ebitda")),1),
                    "gross_margin":   round(safe_float(pd.get("gross_margin"))*100,1),
                    "op_margin":      round(safe_float(pd.get("op_margin"))*100,1),
                    "revenue_growth": round(safe_float(pd.get("revenue_growth"))*100,1),
                }
            except Exception: pass

        return jsonify({"ok":True,"ticker":ticker,"filename":fname,
                        "summary":summary,"dcf":dcf,"peers":peers_out})

    except Exception as e:
        log.error(f"[{ticker}] FAILED {type(e).__name__}: {e}")
        log.error(traceback.format_exc())
        msg = str(e)
        if "yfinance" in msg.lower() or "No module" in msg:
            msg = "yfinance not installed. Check /test for diagnostics."
        elif "CIK not found" in msg or "not found" in msg.lower():
            msg = f"Ticker '{ticker}' not found. Please verify the symbol."
        elif "Connection" in msg or "timeout" in msg.lower() or "Timeout" in msg:
            msg = "Cannot reach Yahoo Finance. Check /test — Render free tier may block outbound requests."
        elif "NoneType" in msg:
            msg = f"Insufficient data for '{ticker}'. Try a major US-listed stock like AAPL or MSFT."
        return jsonify({"error": msg}), 500


@app.route("/api/download/<filename>")
def download(filename):
    safe_name = Path(filename).name
    if not safe_name.endswith(".xlsx"):
        return jsonify({"error":"Invalid file"}), 400
    path = Path("/tmp/stockapp") / safe_name
    if not path.exists():
        return jsonify({"error":"File not found. Please re-run the analysis."}), 404
    return send_file(str(path), as_attachment=True, download_name=safe_name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.errorhandler(404)
def not_found(e):
    return jsonify({"error":"Not found"}), 404

@app.errorhandler(500)
def server_error(e):
    log.error(f"500: {e}\n{traceback.format_exc()}")
    return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    log.info(f"Starting on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
