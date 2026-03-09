"""
Exit Live Trades via IB Gateway with optional Open → MOC changes in Live_Trade_Info.xlsx

Same as Exit_GW plus: before reading exit types, asks if any symbols should have their
order type changed from "Open" to "MOC". If yes, lists symbols, you enter which ones
(comma-space separated, up to 10); script writes "MOC" to column D for those rows,
saves the workbook, then reads column D and sends MOC/MKT orders accordingly.

- Reads/writes Live_Trade_Info.xlsx (sheet "Prices"), columns A–D via xlwings.
- Column D: "Open" → MKT, else → MOC. You can change Open → MOC before sending.
- Still prompts "Send live exit orders? (y/n)" before connecting and sending.

Prerequisites: IB Gateway running (API enabled), pip install ib_insync xlwings.
Close Live_Trade_Info.xlsx in Excel before running.
"""

import os
import sys
import asyncio
from typing import List, Tuple

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

IB_IMPORT_ERROR = None
try:
    from ib_insync import IB, Stock, Order
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
LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
LIVE_INFO_SHEET = "Prices"

IB_HOST = "127.0.0.1"
IB_PORT = 4001
IB_CLIENT_ID = 4
IB_ACCOUNT = "U24159961"
DEFAULT_EXCHANGE = "SMART"
DEFAULT_CURRENCY = "USD"
DRY_RUN = False

MAX_SYMBOLS_TO_CHANGE = 10


def normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    s = str(direction_cell).strip().lower()
    if s in ("long", "short"):
        return s
    return s


def _exit_type_cell_to_order_type(cell_value) -> str:
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def _get_symbols_from_sheet(sheet) -> List[str]:
    """Return list of non-empty tickers from column A (row 2 onward), order preserved."""
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000
    symbols: List[str] = []
    for row in range(2, max_row + 1):
        cell = sheet.range(f"A{row}").value
        if cell is None or str(cell).strip() == "":
            continue
        symbols.append(str(cell).strip().upper())
    return symbols


def _set_exit_type_to_moc(sheet, tickers: List[str]) -> int:
    """
    For each ticker in tickers, find row(s) in column A where value matches (case-insensitive)
    and set column D to "MOC". Returns number of cells updated.
    """
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000
    ticker_set = {t.upper() for t in tickers}
    updated = 0
    for row in range(2, max_row + 1):
        cell = sheet.range(f"A{row}").value
        if cell is None or str(cell).strip() == "":
            continue
        if str(cell).strip().upper() in ticker_set:
            sheet.range(f"D{row}").value = "MOC"
            updated += 1
    return updated


def _parse_moc_input(user_input: str) -> List[str]:
    """Parse 'AAPL, WMT, MSFT' into ['AAPL','WMT','MSFT'], max 10, strip and uppercase."""
    parts = [p.strip().upper() for p in user_input.split(",") if p.strip()]
    return parts[:MAX_SYMBOLS_TO_CHANGE]


def read_exit_trade_info(sheet):
    """
    Read columns A–D from sheet. Column D = IB Exit.
    Returns (exits, all_are_open): exits = [(ticker, action, size, order_type), ...],
    all_are_open = True iff every row's column D is "Open".
    """
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000

    exits: List[Tuple[str, str, int, str]] = []
    all_are_open = True

    for row in range(2, max_row + 1):
        ticker_cell = sheet.range(f"A{row}").value
        direction_cell = sheet.range(f"B{row}").value
        size_cell = sheet.range(f"C{row}").value
        exit_type_cell = sheet.range(f"D{row}").value

        if ticker_cell is None or str(ticker_cell).strip() == "":
            continue

        ticker = str(ticker_cell).strip().upper()
        direction_norm = normalize_direction(direction_cell)

        if direction_norm not in ("long", "short"):
            print(f"Row {row}: invalid direction '{direction_cell}' for ticker {ticker}; skipping.")
            continue

        try:
            size = int(size_cell)
        except (TypeError, ValueError):
            print(f"Row {row}: invalid share size '{size_cell}' for ticker {ticker}; skipping.")
            continue

        if size <= 0:
            print(f"Row {row}: non-positive share size {size} for ticker {ticker}; skipping.")
            continue

        order_type = _exit_type_cell_to_order_type(exit_type_cell)
        if order_type != "MKT":
            all_are_open = False
        action = "SELL" if direction_norm == "long" else "BUY"
        exits.append((ticker, action, size, order_type))

    return exits, all_are_open


def connect_ib() -> IB:
    if IB is None:
        raise ImportError(
            "Could not import ib_insync or its dependencies. "
            f"Details:\n    {IB_IMPORT_ERROR}\n\n"
            "Try: python -m pip install --upgrade ib_insync eventkit nest-asyncio numpy"
        )
    ib = IB()
    print(f"Connecting to IB Gateway at {IB_HOST}:{IB_PORT} with clientId={IB_CLIENT_ID} ...")
    ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
    print("Connected to IB Gateway.")
    return ib


def place_exit_orders_ib(ib: IB, exits: List[Tuple[str, str, int, str]]) -> None:
    if not exits:
        print("No exit orders to place.")
        return

    print("\nPlanned exit orders (close/cover):")
    for ticker, action, size, order_type in exits:
        print(f"  {action} {size} {ticker}  [{order_type}]")

    if DRY_RUN:
        print("\nDRY_RUN is True: no orders will be sent. Set DRY_RUN = False to send exit orders.")
        return

    print("\nPlacing exit orders...")
    for ticker, action, size, order_type in exits:
        contract = Stock(ticker, DEFAULT_EXCHANGE, DEFAULT_CURRENCY)
        order = Order(
            action=action,
            orderType=order_type,
            totalQuantity=size,
            tif="DAY",
        )
        if IB_ACCOUNT:
            order.account = IB_ACCOUNT
        trade = ib.placeOrder(contract, order)
        print(f"Submitted {action} {size} {ticker} ({order_type}), orderId={trade.order.orderId}")

    ib.sleep(2)
    print("\nOrder statuses:")
    for t in ib.trades():
        print(
            f"  orderId={t.order.orderId} status={t.orderStatus.status} "
            f"filled={t.orderStatus.filled} remaining={t.orderStatus.remaining}"
        )


def main():
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return
    if not os.path.exists(LIVE_INFO_FILE):
        print(f"Live trade info file not found: {LIVE_INFO_FILE}")
        return

    app = None
    wb = None
    try:
        app = xw.App(visible=False)
        try:
            wb = app.books.open(os.path.abspath(LIVE_INFO_FILE))
        except Exception:
            print("Please close Live_Trade_Info")
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass
                app = None
            return
        try:
            sheet = wb.sheets[LIVE_INFO_SHEET]
        except Exception:
            print(f"Sheet '{LIVE_INFO_SHEET}' not found in {LIVE_INFO_FILE}.")
            wb.close()
            return

        # Get symbols for optional Open → MOC prompt
        symbols = _get_symbols_from_sheet(sheet)
        if not symbols:
            print("No symbols found in Live_Trade_Info; nothing to do.")
            wb.close()
            sys.exit(0)  # Clean exit for scheduled runs with no trades

        symbols_str = ", ".join(symbols)
        reply = input(
            "\nAre there any symbols that need their order type to change from \"Open\" to \"MOC\"? (y/n): "
        ).strip().lower()
        if reply in ("y", "yes"):
            which = input(
                f"Which ones to change to MOC ({symbols_str})? "
            ).strip()
            tickers_to_moc = _parse_moc_input(which)
            if tickers_to_moc:
                n = _set_exit_type_to_moc(sheet, tickers_to_moc)
                wb.save()
                print(f"Updated {n} row(s) to MOC for: {', '.join(tickers_to_moc)}.")

        exits, all_are_open = read_exit_trade_info(sheet)

        if not exits:
            print("No valid rows in Live_Trade_Info; nothing to exit.")
            wb.close()
            sys.exit(0)  # Clean exit for scheduled runs with no trades

        if not all_are_open:
            reply = input("\nSome IB exit types aren't exits at the open, send anyway? (y/n): ").strip().lower()
            if reply not in ("y", "yes"):
                print("Exiting without sending orders.")
                wb.close()
                return

        reply = input("\nSend live exit orders? (y/n): ").strip().lower()
        if reply not in ("y", "yes"):
            print("Exiting without sending orders.")
            wb.close()
            return

        try:
            ib = connect_ib()
        except Exception as e:
            print(f"Failed to connect to IB Gateway: {e}")
            wb.close()
            return

        try:
            place_exit_orders_ib(ib, exits)
        finally:
            print("Disconnecting from IB Gateway...")
            ib.disconnect()
            print("Disconnected.")

        wb.save()
        wb.close()
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
