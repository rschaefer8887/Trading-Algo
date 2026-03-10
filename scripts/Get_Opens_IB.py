"""
Get Opens IB — Opening prices from Interactive Brokers Gateway into Latest Earnings

Same workflow as Get_Opens_Auto but uses IB Gateway (ib_insync) instead of yfinance
to fetch opening prices. Reads the Latest Earnings workbook (Trades sheet):
  - Column K: first "O" starts the range, "0" stops (only one O and one 0 through row 550).
  - Column A: ticker per row. Opening price is written to column T.

IB Gateway must be running with API enabled (default port 4001 live, 4002 paper).
The workbook is opened and saved via xlwings. Close it in Excel before running.
"""

import os
import sys
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO

try:
    import asyncio
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

IB_IMPORT_ERROR = None
try:
    from ib_insync import IB, Stock
except Exception as e:
    IB = None  # type: ignore[assignment]
    IB_IMPORT_ERROR = e

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
HEADER_ROW = 3
COL_TICKER = "A"
COL_FLAG = "K"
COL_OPENING_PRICE = "T"
CHECK_END_ROW = 550

# IB Gateway (same port as other Gateway scripts; use a distinct client ID)
IB_HOST = "127.0.0.1"
IB_PORT = 4001  # 4001 = live, 4002 = paper
IB_CLIENT_ID = 5  # Distinct from Open_Trades_GW (3) and Exit_GW (4)
IB_EXCHANGE = "SMART"
IB_CURRENCY = "USD"

# Historical bar request for daily open
IB_DURATION = "5 D"
IB_BAR_SIZE = "1 day"
IB_WHAT_TO_SHOW = "TRADES"
IB_USE_RTH = True


def _is_start_flag(cell_value) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "O"


def _is_stop_flag(cell_value) -> bool:
    if cell_value is None:
        return False
    if isinstance(cell_value, (int, float)):
        return cell_value == 0
    return str(cell_value).strip() == "0"


def _normalize_ticker(cell_value):
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _ticker_for_ib(ticker: str) -> str:
    """Symbol for IB contract (use dot for BRK.B, etc.)."""
    if not ticker:
        return ""
    return ticker.replace("-", ".")


def _gateway_connected() -> bool:
    """Try to connect to IB Gateway; return True if connected. Suppresses library output on failure."""
    if IB is None:
        return False
    ib = IB()
    try:
        with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
            ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
    except Exception:
        return False
    try:
        ib.disconnect()
    except Exception:
        pass
    return True


def _fetch_opens_via_ib(tickers: list[str]) -> dict[str, float | None]:
    """Fetch latest daily open for each ticker from IB Gateway. Returns dict ticker -> price."""
    if not tickers or IB is None:
        return {t: None for t in tickers} if tickers else {}
    if IB_IMPORT_ERROR is not None:
        return {t: None for t in tickers}

    unique = list(dict.fromkeys(tickers))
    result: dict[str, float | None] = {}
    ib = IB()
    try:
        ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
    except Exception:
        return {t: None for t in unique}

    try:
        for t in unique:
            sym = _ticker_for_ib(t)
            try:
                contract = Stock(sym, IB_EXCHANGE, IB_CURRENCY)
                bars = ib.reqHistoricalData(
                    contract,
                    endDateTime="",
                    durationStr=IB_DURATION,
                    barSizeSetting=IB_BAR_SIZE,
                    whatToShow=IB_WHAT_TO_SHOW,
                    useRTH=IB_USE_RTH,
                )
                if bars:
                    open_val = getattr(bars[-1], "open", None)
                    result[t] = round(float(open_val), 2) if open_val is not None else None
                else:
                    result[t] = None
            except Exception as e:
                print(f"  Warning: IB no open for {t}: {e}")
                result[t] = None
    finally:
        try:
            ib.disconnect()
        except Exception:
            pass

    return result


def main():
    if IB is None or IB_IMPORT_ERROR is not None:
        print("ib_insync is required. Install with: pip install ib_insync")
        if IB_IMPORT_ERROR:
            print(f"  Error: {IB_IMPORT_ERROR}")
        return
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return
    if not os.path.exists(SOURCE_FILE):
        print(f"Source file not found: {SOURCE_FILE}")
        return

    if not _gateway_connected():
        print("Gateway is not open, please open the gateway and try again.")
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

        # Ensure exactly one O and one 0 in column K through row 550
        count_o = 0
        count_zero = 0
        end_check = min(max_row, CHECK_END_ROW)
        for row in range(start_row, end_check + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_start_flag(cell_val):
                count_o += 1
            elif _is_stop_flag(cell_val):
                count_zero += 1
        if count_o != 1 or count_zero != 1:
            print("One clean range is not selected, please clean up your open range and try again.")
            wb.close()
            sys.exit(0)  # Clean exit for scheduled runs with no valid range

        first_o_row = None
        for row in range(start_row, max_row + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_start_flag(cell_val):
                first_o_row = row
                break

        if first_o_row is None:
            print("No 'O' flag found in column K. Nothing to process.")
            wb.close()
            sys.exit(0)  # Clean exit for scheduled runs with no O in column K

        to_process: list[tuple[int, str]] = []
        for row in range(first_o_row, max_row + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_stop_flag(cell_val):
                break
            ticker_raw = sheet.range(f"{COL_TICKER}{row}").value
            ticker = _normalize_ticker(ticker_raw)
            if not ticker:
                continue
            to_process.append((row, ticker))

        if not to_process:
            print("No tickers found in rows between 'O' and '0' in column K.")
            wb.close()
            sys.exit(0)  # Clean exit for scheduled runs with no tickers in range

        tickers_to_fetch = [t for _, t in to_process]
        print(f"Fetching opening prices for {len(tickers_to_fetch)} ticker(s) from IB Gateway...")
        prices = _fetch_opens_via_ib(tickers_to_fetch)

        written: list[str] = []
        for row, ticker in to_process:
            price = prices.get(ticker)
            if price is not None:
                sheet.range(f"{COL_OPENING_PRICE}{row}").value = price
                written.append(ticker)
            else:
                print(f"  Warning: no opening price for {ticker} (row {row})")

        wb.save()
        wb.close()
        print("Opening prices written to Latest Earnings (saved via Excel).")
        print(f"\nTickers written to column {COL_OPENING_PRICE}: {', '.join(written)}")
    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()
