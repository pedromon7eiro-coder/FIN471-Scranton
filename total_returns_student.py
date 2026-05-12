## ============================================================
##  Daily Total Return Downloader
##  Source: Yahoo Finance
## ------------------------------------------------------------
##  Professor Pedro Monteiro
##  FIN 471 — University of Scranton
## ============================================================
##
##  WHAT THIS SCRIPT GENERATES
##  --------------------------
##  An Excel workbook (daily_total_returns.xlsx) with one sheet
##  per ticker, plus a combined Summary sheet. Each sheet contains:
##
##    • Date          — trading day (YYYY-MM-DD)
##    • Close         — actual raw closing price investors transact at
##    • Dividends     — cash dividend paid on that date (0 on most days)
##    • Total Return  — (Close_t + Dividends_t - Close_{t-1}) / Close_{t-1}
##                      true buy-and-hold daily total return, including
##                      any dividend received that day
##
##  The Summary sheet shows all tickers side-by-side for easy comparison.
##
##  NOTE ON PRICES
##  --------------
##  Prices are raw market closing prices (what an investor actually paid
##  or received), NOT backward-adjusted prices. Dividends are captured
##  explicitly in the Dividends column and included in the Total Return
##  formula above, so no return information is lost.
##
##  NOTE ON THE FIRST ROW
##  ---------------------
##  The first trading day in the date range has no prior closing price,
##  so its Total Return is left blank (N/A). All subsequent rows are complete.
##
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
OUTPUT_FILE = "daily_total_returns.xlsx"

## ============================================================
##  DO NOT EDIT BELOW THIS LINE
## ============================================================
!pip install yfinance openpyxl -q

import yfinance as yf
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ── palette ────────────────────────────────────────────────────
NAVY      = "0D1B2A"
STEEL     = "1B3A5C"
GOLD      = "C9A84C"
WHITE     = "FFFFFF"
LIGHT_BG  = "F4F6FA"
ALT_ROW   = "EAF0FB"
BORDER_C  = "C5CFE0"
GREEN_POS = "1D7A4A"
RED_NEG   = "C0392B"
DIV_COLOR = "7B5EA7"

def side(color=BORDER_C, style="thin"):
    return Side(style=style, color=color)

THIN_BORDER = Border(left=side(), right=side(), top=side(), bottom=side())
HEADER_BORDER = Border(
    left=side(NAVY, "medium"), right=side(NAVY, "medium"),
    top=side(NAVY, "medium"), bottom=side(GOLD, "medium")
)

def title_font(sz=11, bold=True, color=WHITE):
    return Font(name="Calibri", size=sz, bold=bold, color=color)

def body_font(sz=10, bold=False, color="000000"):
    return Font(name="Calibri", size=sz, bold=bold, color=color)

def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def right_align():
    return Alignment(horizontal="right", vertical="center")

# ── download data ───────────────────────────────────────────────
end = datetime.today().strftime("%Y-%m-%d") if END_DATE == "today" else END_DATE
ticker_data = {}

for ticker, name in TICKERS.items():
    print(f"Downloading {ticker} ({name})...")
    raw = yf.download(
        ticker,
        start=START_DATE,
        end=end,
        auto_adjust=False,
        actions=True,
        progress=False,
    )
    if raw.empty:
        print(f"  WARNING: No data found for {ticker}.")
        continue

    def get_col(df, col):
        if isinstance(df.columns, pd.MultiIndex):
            return df[(col, ticker)].squeeze()
        return df[col].squeeze()

    close     = get_col(raw, "Close").round(2)
    dividends = get_col(raw, "Dividends")
    total_ret = (close + dividends - close.shift(1)) / close.shift(1)

    df = pd.DataFrame({
        "Date":         close.index.strftime("%Y-%m-%d"),
        "Close":        close.values,
        "Dividends":    dividends.values,
        "Total_Return": total_ret.values,
    }).iloc[1:].reset_index(drop=True)   # drop first row (no prior price)

    ticker_data[ticker] = (name, df)

if not ticker_data:
    print("No data downloaded. Check your ticker symbols.")
    raise SystemExit

# ── build workbook ──────────────────────────────────────────────
wb = Workbook()
wb.remove(wb.active)   # remove default blank sheet

def write_sheet(wb, sheet_name, ticker, name, df):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    # ── banner (rows 1-2) ──────────────────────────────────────
    ws.merge_cells("A1:D2")
    banner = ws["A1"]
    banner.value = f"{ticker}  ·  {name}"
    banner.font  = Font(name="Calibri", size=16, bold=True, color=WHITE)
    banner.fill  = fill(NAVY)
    banner.alignment = center()

    # ── sub-banner (row 3) ────────────────────────────────────
    ws.merge_cells("A3:D3")
    sub = ws["A3"]
    sub.value = f"Daily Total Return  |  {df['Date'].iloc[0]}  →  {df['Date'].iloc[-1]}  |  {len(df):,} trading days"
    sub.font  = Font(name="Calibri", size=10, italic=True, color="D0D8E8")
    sub.fill  = fill(STEEL)
    sub.alignment = center()

    # ── column headers (row 4) ────────────────────────────────
    headers = ["Date", "Close ($)", "Dividends ($)", "Total Return (%)"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font      = title_font(10)
        cell.fill      = fill(STEEL)
        cell.alignment = center()
        cell.border    = HEADER_BORDER

    # ── data rows ─────────────────────────────────────────────
    for row_idx, row in df.iterrows():
        excel_row = row_idx + 5
        bg = ALT_ROW if row_idx % 2 == 0 else LIGHT_BG

        # Date
        c = ws.cell(row=excel_row, column=1, value=row["Date"])
        c.font = body_font(color="2C3E50"); c.fill = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

        # Close
        c = ws.cell(row=excel_row, column=2, value=row["Close"])
        c.font = body_font(); c.fill = fill(bg)
        c.number_format = '#,##0.00'; c.alignment = right_align()
        c.border = THIN_BORDER

        # Dividends
        div_val = row["Dividends"]
        c = ws.cell(row=excel_row, column=3, value=div_val if div_val > 0 else None)
        c.font = Font(name="Calibri", size=10, bold=div_val > 0, color=DIV_COLOR if div_val > 0 else "AAAAAA")
        c.fill = fill(bg)
        c.number_format = '#,##0.0000'; c.alignment = right_align()
        c.border = THIN_BORDER

        # Total Return
        ret_val = row["Total_Return"]
        if pd.notna(ret_val):
            c = ws.cell(row=excel_row, column=4, value=ret_val)
            color = GREEN_POS if ret_val >= 0 else RED_NEG
            c.font = Font(name="Calibri", size=10, bold=True, color=color)
            c.number_format = '0.0000%'; c.alignment = right_align()
        else:
            c = ws.cell(row=excel_row, column=4, value="N/A")
            c.font = body_font(color="AAAAAA"); c.alignment = right_align()
        c.fill = fill(bg); c.border = THIN_BORDER

    # ── column widths ─────────────────────────────────────────
    for col, width in zip(["A","B","C","D"], [14, 13, 16, 18]):
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 20

    # ── color-scale conditional formatting on Total Return ────
    last_row = len(df) + 4
    ws.conditional_formatting.add(
        f"D5:D{last_row}",
        ColorScaleRule(
            start_type="min",  start_color="FFCCCC",
            mid_type="num",    mid_value=0, mid_color="FFFFFF",
            end_type="max",    end_color="CCFFDD",
        )
    )

    print(f"  ✓ Sheet '{sheet_name}' written — {len(df):,} rows")
    return ws

# ── individual ticker sheets ────────────────────────────────────
for ticker, (name, df) in ticker_data.items():
    write_sheet(wb, ticker, ticker, name, df)

# ── summary sheet ───────────────────────────────────────────────
ws = wb.create_sheet("Summary", 0)
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

all_dates = sorted(set(
    date for _, df in ticker_data.values() for date in df["Date"]
))
tickers_list = list(ticker_data.keys())

# banner
num_cols = 1 + len(tickers_list) * 2
ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=num_cols)
b = ws.cell(row=1, column=1)
b.value = "FIN 471  ·  Daily Total Return Summary  ·  University of Scranton"
b.font  = Font(name="Calibri", size=15, bold=True, color=WHITE)
b.fill  = fill(NAVY); b.alignment = center()

ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=num_cols)
s = ws.cell(row=3, column=1)
s.value = f"Professor Pedro Monteiro  |  {all_dates[0]}  →  {all_dates[-1]}  |  Source: Yahoo Finance"
s.font  = Font(name="Calibri", size=10, italic=True, color="D0D8E8")
s.fill  = fill(STEEL); s.alignment = center()

# ticker group headers (row 4)
ws.cell(row=4, column=1, value="Date").font = title_font(10)
ws["A4"].fill = fill(STEEL); ws["A4"].alignment = center()
ws["A4"].border = HEADER_BORDER

col = 2
ticker_col_map = {}
for ticker in tickers_list:
    name = ticker_data[ticker][0]
    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
    cell = ws.cell(row=4, column=col, value=f"{ticker} — {name}")
    cell.font = title_font(10); cell.fill = fill(STEEL)
    cell.alignment = center(); cell.border = HEADER_BORDER
    ticker_col_map[ticker] = col
    col += 2

# sub-headers (row 5) — Close and Total Return per ticker
ws.cell(row=5, column=1, value="Date").font = title_font(9)
ws["A5"].fill = fill(STEEL); ws["A5"].alignment = center()
ws["A5"].border = HEADER_BORDER

for ticker in tickers_list:
    c = ticker_col_map[ticker]
    for offset, label in enumerate(["Close ($)", "Total Return (%)"]):
        cell = ws.cell(row=5, column=c+offset, value=label)
        cell.font = title_font(9); cell.fill = fill(STEEL)
        cell.alignment = center(); cell.border = HEADER_BORDER

ws.freeze_panes = "A6"

# build lookup for fast access
data_lookup = {
    ticker: df.set_index("Date") for ticker, (_, df) in ticker_data.items()
}

for row_idx, date in enumerate(all_dates):
    excel_row = row_idx + 6
    bg = ALT_ROW if row_idx % 2 == 0 else LIGHT_BG

    c = ws.cell(row=excel_row, column=1, value=date)
    c.font = body_font(color="2C3E50"); c.fill = fill(bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = THIN_BORDER

    for ticker in tickers_list:
        col = ticker_col_map[ticker]
        df_t = data_lookup[ticker]

        if date in df_t.index:
            row_data = df_t.loc[date]
            # Close
            cl = ws.cell(row=excel_row, column=col, value=row_data["Close"])
            cl.font = body_font(); cl.fill = fill(bg)
            cl.number_format = '#,##0.00'; cl.alignment = right_align()
            cl.border = THIN_BORDER
            # Total Return
            ret_val = row_data["Total_Return"]
            if pd.notna(ret_val):
                rc = ws.cell(row=excel_row, column=col+1, value=ret_val)
                color = GREEN_POS if ret_val >= 0 else RED_NEG
                rc.font = Font(name="Calibri", size=10, bold=True, color=color)
                rc.number_format = '0.0000%'
            else:
                rc = ws.cell(row=excel_row, column=col+1, value="N/A")
                rc.font = body_font(color="AAAAAA")
            rc.fill = fill(bg); rc.alignment = right_align(); rc.border = THIN_BORDER
        else:
            for offset in range(2):
                ec = ws.cell(row=excel_row, column=col+offset, value="—")
                ec.font = body_font(color="CCCCCC"); ec.fill = fill(bg)
                ec.alignment = center(); ec.border = THIN_BORDER

# column widths for summary
ws.column_dimensions["A"].width = 14
col = 2
for ticker in tickers_list:
    ws.column_dimensions[get_column_letter(col)].width   = 13
    ws.column_dimensions[get_column_letter(col+1)].width = 17
    col += 2

ws.row_dimensions[1].height = 28
ws.row_dimensions[3].height = 18
ws.row_dimensions[4].height = 20
ws.row_dimensions[5].height = 18

# conditional formatting on each Total Return column in summary
last_row = len(all_dates) + 5
for ticker in tickers_list:
    col = ticker_col_map[ticker] + 1
    col_letter = get_column_letter(col)
    ws.conditional_formatting.add(
        f"{col_letter}6:{col_letter}{last_row}",
        ColorScaleRule(
            start_type="min",  start_color="FFCCCC",
            mid_type="num",    mid_value=0, mid_color="FFFFFF",
            end_type="max",    end_color="CCFFDD",
        )
    )

print(f"  ✓ Summary sheet written — {len(all_dates):,} trading days × {len(tickers_list)} tickers")

wb.save(OUTPUT_FILE)
print(f"\nComplete. Workbook saved as '{OUTPUT_FILE}'.")
print(f"Sheets: Summary + {list(ticker_data.keys())}")

from google.colab import files
files.download(OUTPUT_FILE)
