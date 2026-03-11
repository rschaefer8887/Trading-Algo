"""
Exit Live Trades to Interactive Brokers via IB Gateway based on Live_Trade_Info.xlsx

This script connects to IB Gateway (not TWS). Use it when Gateway is running for
exits. Gateway must be running with API enabled. Default ports: 4001 = live, 4002 = paper.

Reads tickers, direction, share size, and exit type from Live_Trade_Info.xlsx (columns A–D)
via xlwings; the workbook is saved on exit so you can add writes later.
Column D (header "IB Exit") is written by Stage_Trades_Auto; values map to order type:
- "Open" → Market order (MKT, executes during the session)
- Other   → Market-on-close order (MOC)

Places exit orders in the **opposite** direction to close/cover:
- Entry was LONG  (bought)  → Exit: SELL (same size)
- Entry was SHORT (sold)     → Exit: BUY  (same size)

If any row has an exit type other than "Open", the script prompts:
"Some IB exit types aren't exits at the open, send anyway? (y/n)" — only sends if y.
Then prompts "Send live exit orders? (y/n)" before connecting and sending.

Prerequisites:
- IB Gateway running with API enabled (socket port 4001 live / 4002 paper).
- Python packages: pip install ib_insync xlwings
- Close Live_Trade_Info.xlsx in Excel before running.
"""

import os
import asyncio
from typing import List, Tuple

# Ensure there is an asyncio event loop for ib_insync/eventkit on newer Python versions.
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
# Paths: Excel files live in repo root (Trading_Algo); script is in Old Scripts
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(os.path.dirname(_SCRIPT_DIR))

# ---------------------------------------------------------------------------
# Configuration — IB Gateway (not TWS)
# ---------------------------------------------------------------------------
LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
LIVE_INFO_SHEET = "Prices"
# Column D = IB Exit (written by Stage_Trades_Auto); "Open" → MKT, else → MOC

# IB Gateway API connection (4001 = live, 4002 = paper). Use different client ID than Open_Trades_GW.
IB_HOST = "127.0.0.1"
IB_PORT = 4001  # 4001 = Gateway live, 4002 = Gateway paper
IB_CLIENT_ID = 4  # Different from Gateway entry script (3) so both can run if needed
IB_ACCOUNT = "U1867866"

DEFAULT_EXCHANGE = "SMART"
DEFAULT_CURRENCY = "USD"

# When True, only print planned exit orders; do not send them.
DRY_RUN = False


def normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    s = str(direction_cell).strip().lower()
    if s in ("long", "short"):
        return s
    return s


def _exit_type_cell_to_order_type(cell_value) -> str:
    """'Open' (case-insensitive) -> 'MKT', else -> 'MOC'."""
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def read_exit_trade_info(sheet):
    """
    Read Live_Trade_Info from the given xlwings sheet (columns A–D). Column D = IB Exit.
    Build exit orders: long → SELL, short → BUY, same size; order type from column D.
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
    """Place exit orders (MOC or MKT per row) to close/cover positions."""
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
            orderType=order_type,  # "MOC" or "MKT"
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

        exits, all_are_open = read_exit_trade_info(sheet)

        if not exits:
            print("No valid rows in Live_Trade_Info; nothing to exit.")
            wb.close()
            return

        # If any exit type is not "Open", prompt before proceeding
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
