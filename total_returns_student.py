## ============================================================
##  Daily Total Return Downloader
##  Source: Yahoo Finance (dividend-adjusted closing prices)
## ------------------------------------------------------------
##  Professor Pedro Monteiro
##  FIN 471 — University of Scranton
## ============================================================
##  INSTRUCTIONS: Edit only the three parameters below.
##  Do not modify anything beneath the divider line.
## ============================================================

# --------------------------------------------------------------
#  PARAMETER 1: Tickers
#  Format: "TICKER": "Description"
#
#  Reference:
#    Equities   -- SPY (S&P 500), IWM (Russell 2000), QQQ (Nasdaq),
#                  DIA (Dow Jones), VTI (Total US Market)
#    High Yield -- HYG, JNK
#    Corp Bond  -- LQD (Investment Grade)
#    Treasuries -- TLT (20+ yr), IEF (7-10 yr), SHY (1-3 yr)
#    Intl       -- EFA (Developed Markets), EEM (Emerging Markets)
#    REIT       -- VNQ
# --------------------------------------------------------------

TICKERS = {
    "SPY": "S&P 500",
    "IWM": "Russell 2000",
    "HYG": "High Yield Bond",
    "TLT": "Treasury Bond (20+ yr)",
}

# --------------------------------------------------------------
#  PARAMETER 2: Date Range
# --------------------------------------------------------------

START_DATE = "2020-01-01"   # YYYY-MM-DD
END_DATE   = "today"        # "today" or a specific date: "2026-05-09"

# --------------------------------------------------------------
#  PARAMETER 3: Output File Name
# --------------------------------------------------------------

OUTPUT_FILE = "daily_total_returns.csv"

## ============================================================
##  DO NOT EDIT BELOW THIS LINE
## ============================================================

!pip install yfinance -q

import yfinance as yf
import pandas as pd
from datetime import datetime

end = datetime.today().strftime("%Y-%m-%d") if END_DATE == "today" else END_DATE

frames = []
for ticker, name in TICKERS.items():
    print(f"Downloading {ticker} ({name})...")
    raw = yf.download(
        ticker,
        start=START_DATE,
        end=end,
        auto_adjust=True,
        actions=True,
        progress=False,
    )
    if raw.empty:
        print(f"  WARNING: No data found for {ticker}. Please verify the ticker symbol.")
        continue

    close = raw["Close"].squeeze().rename(f"{ticker}_Adj_Close")
    ret   = (close.pct_change() * 100).round(4).apply(lambda x: f"{x:.4f}%" if not __import__("math").isnan(x) else "").rename(f"{ticker}_Daily_Return_%")
    frames.extend([close, ret])

if not frames:
    print("No data was downloaded. Please review your ticker symbols and try again.")
else:
    df = pd.concat(frames, axis=1)
    df.index = pd.to_datetime(df.index).strftime("%Y-%m-%d")
    df.index.name = "Date"
    price_cols = [c for c in df.columns if "Adj_Close" in c]
    df = df.dropna(subset=price_cols, how="all")

    # Format adjusted close prices to 2 decimal places
    for col in df.columns:
        if "Adj_Close" in col:
            df[col] = df[col].round(2)

    df.to_csv(OUTPUT_FILE)
    print(f"\nComplete. {len(df)} trading days saved to '{OUTPUT_FILE}'.")
    print(f"Date range : {df.index[0]}  to  {df.index[-1]}")
    print(f"Columns    : {list(df.columns)}")

    print(f"\nPreview (first 5 rows):")
    print(df.head(5).to_string())

    print(f"\nSaving {OUTPUT_FILE} to your local machine...")
    from google.colab import files
    files.download(OUTPUT_FILE)
