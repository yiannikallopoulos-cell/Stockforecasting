# Equilens — Stock Analysis Web App

A full-stack financial analysis web app that runs a DCF model on any public
company, generates a downloadable Excel workbook, and produces a Buy/Hold/Sell
rating by comparing the DCF implied price to the current market price.

## Setup

### 1. Install dependencies

Open a terminal in PyCharm (or your system terminal), navigate to this folder,
and run:

```bash
pip install -r requirements.txt
```

### 2. Copy your model file

Make sure `model.py` (your `Forecasting_v2.py`) is in the same folder as
`app.py`. It should already be there if you downloaded the full package.

### 3. Run the app

```bash
python app.py
```

Then open your browser to: **http://localhost:5000**

---

## How it works

1. **Enter a ticker** (e.g. META, AAPL, NVDA) and select projection years
2. The app fetches data from **Yahoo Finance** (via yfinance) and **SEC EDGAR**
3. A full **DCF model** is computed — projected FCFs, terminal value, equity value
4. The **implied share price** is compared to the current market price
5. A **rating** is assigned based on the % difference:
   - **Strong Buy** — DCF implies ≥ +30% upside
   - **Buy** — DCF implies +10% to +30% upside
   - **Hold** — DCF within ±10% of current price
   - **Sell** — DCF implies -10% to -25% downside
   - **Strong Sell** — DCF implies > -25% downside
6. A full **Excel workbook** is available to download (6 sheets: Cover,
   Assumptions, Income Statement, Balance Sheet, Cash Flow, Valuation)

---

## File structure

```
stockapp/
├── app.py              ← Flask backend + DCF rating logic
├── model.py            ← Financial model (your Forecasting_v2.py)
├── requirements.txt    ← Python dependencies
├── templates/
│   └── index.html      ← Frontend (single-page app)
└── README.md
```

---

## Notes

- Results are cached in `/tmp/stockapp/` for download. Files are not persisted
  between server restarts.
- The DCF model is a simplified equity DCF — it does not account for all
  company-specific factors. Always do your own due diligence.
- yfinance may occasionally fail for certain tickers or during Yahoo Finance
  maintenance windows. If you get an error, wait 30 seconds and retry.
