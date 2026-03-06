"""
Get CP Auto — Automated closing prices into Latest Earnings by column L flags

Reads the Latest Earnings workbook (Trades sheet). Column L contains flags:
  - First "1": start; write closing prices to column AS until the next 1.
  - Next "1": switch to column AL until the next 1.
  - Next "1": switch to column S until a "0" flag.
  - "0": stop; do not process that row.

For each processed row: ticker from column A, latest close from yfinance,
written to the current target column (AS, AL, or S). Prints a summary of
tickers written to each column.

The Earnings file is opened and saved via xlwings (Excel) so that external
links and other Excel features are preserved. Close the workbook in Excel
before running.
"""

import os

import yfinance as yf

try:
    import xlwings as xw
except ImportError:
    xw = None

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Document.xlsx")
SOURCE_SHEET = "Trades"
HEADER_ROW = 3  # Data starts the row after headers (e.g. row 4)
COL_TICKER = "A"
COL_FLAG = "L"
COL_AS = "AS"
COL_AL = "AL"
COL_S = "S"


def _flag_value(cell_value):
    """Normalize cell to '0', '1', or None/other. Accepts numeric 1/0 from Excel (e.g. 1.0, 0.0)."""
    if cell_value is None:
        return None
    if isinstance(cell_value, (int, float)):
        if cell_value == 1:
            return "1"
        if cell_value == 0:
            return "0"
        return None
    s = str(cell_value).strip()
    if s in ("0", "1"):
        return s
    return None


def _normalize_ticker(cell_value):
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _fetch_prices_batch(tickers: list[str]) -> dict[str, float | None]:
    """Fetch latest closing price for each ticker via yfinance (batch). Returns dict ticker -> price."""
    if not tickers:
        return {}
    unique = list(dict.fromkeys(tickers))
    try:
        data = yf.download(
            unique, period="5d", group_by="ticker", auto_adjust=True, progress=False, threads=True
        )
    except Exception:
        return {t: None for t in unique}
    if data.empty:
        return {t: None for t in unique}
    result = {}
    for t in unique:
        try:
            if len(unique) == 1 and "Close" in data.columns:
                close = data["Close"]
            elif hasattr(data.columns, "get_level_values") and t in data.columns.get_level_values(0):
                close = data[t]["Close"]
            else:
                result[t] = None
                continue
            result[t] = float(close.iloc[-1])
        except Exception:
            result[t] = None
    return result


def main():
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return
    if not os.path.exists(SOURCE_FILE):
        print(f"Source file not found: {SOURCE_FILE}")
        return

    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(os.path.abspath(SOURCE_FILE))
        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"Sheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
            wb.close()
            return

        try:
            max_row = sheet.used_range.last_cell.row
        except Exception:
            max_row = 2000
        start_row = HEADER_ROW + 1

        zone = 0
        target_col_letter = None
        to_process: list[tuple[int, str, str]] = []
        tickers_to_fetch: list[str] = []

        for row in range(start_row, max_row + 1):
            flag_cell = sheet.range(f"{COL_FLAG}{row}").value
            flag = _flag_value(flag_cell)

            if flag == "0":
                break

            if flag == "1":
                zone += 1
                if zone == 1:
                    target_col_letter = COL_AS
                elif zone == 2:
                    target_col_letter = COL_AL
                else:
                    target_col_letter = COL_S

            if zone == 0:
                continue

            ticker_raw = sheet.range(f"{COL_TICKER}{row}").value
            ticker = _normalize_ticker(ticker_raw)
            if not ticker:
                continue

            to_process.append((row, ticker, target_col_letter))
            tickers_to_fetch.append(ticker)

        if not to_process:
            print("No rows to process (no tickers under flag blocks in column L).")
            wb.close()
            return

        print(f"Fetching closing prices for {len(tickers_to_fetch)} ticker(s)...")
        prices = _fetch_prices_batch(tickers_to_fetch)
        tickers_by_column: dict[str, list[str]] = {"AS": [], "AL": [], "S": []}

        for row, ticker, target_col_letter in to_process:
            price = prices.get(ticker)
            if price is not None:
                sheet.range(f"{target_col_letter}{row}").value = price
                tickers_by_column[target_col_letter].append(ticker)
            else:
                print(f"  Warning: no price for {ticker} (row {row})")

        wb.save()
        wb.close()
        print("Closing prices written to Latest Earnings (saved via Excel).")

        print("\nTickers written by column:")
        for col in ("AS", "AL", "S"):
            tickers = tickers_by_column[col]
            if tickers:
                print(f"  Column {col}: {', '.join(tickers)}")
            else:
                print(f"  Column {col}: (none)")
    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()
