"""
Exit_ToS — Exit live trades via Schwab (ToS) based on Live_Trade_Info.xlsx

Logic mirrors Exit_IB_via_GW, but:
- Sends exit orders to Schwab instead of IB Gateway.
- Uses column E ("ToS Exit") for Schwab exit type:
    - "Open" → MARKET (execute during the session)
    - Anything else (including "MOC" from Stage_Trades_Auto) → MARKET_ON_CLOSE

Workbook shape (sheet "Prices" in Live_Trade_Info.xlsx):
- Column A: Ticker
- Column B: Direction ("long" / "short")
- Column C: Share Size
- Column D: IB Exit (used by IB exit scripts, ignored here)
- Column E: ToS Exit (used by this script only)

Flow:
- Reads/writes Live_Trade_Info.xlsx via xlwings; prompts once:
    "Send live exit orders to Schwab? (y/n):"
- For each valid row:
    - LONG  -> action SELL   (close long)
    - SHORT -> action BUY_TO_COVER (close short)
- Order type per row:
    - E == "Open" → Schwab OrderType.MARKET
    - else        → Schwab OrderType.MARKET_ON_CLOSE

Prerequisites:
- schwab-py and xlwings installed:
    python -m pip install --upgrade schwab-py xlwings
- Schwab_Auth.py configured and run at least once to complete OAuth.
- Close Live_Trade_Info.xlsx in Excel before running.
"""

import os
from typing import List, Tuple

import asyncio

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

try:
    import xlwings as xw
except ImportError:
    xw = None

from Schwab_Auth import create_client

SCHWAB_IMPORT_ERROR = None
try:
    from schwab.orders.equities import (
        equity_sell_market,
        equity_buy_to_cover_market,
    )
    from schwab.orders.common import OrderType, EquityInstruction
    from schwab.orders.generic import OrderBuilder
except Exception as e:  # pragma: no cover - import-time failure
    equity_sell_market = None  # type: ignore[assignment]
    equity_buy_to_cover_market = None  # type: ignore[assignment]
    OrderType = None  # type: ignore[assignment]
    EquityInstruction = None  # type: ignore[assignment]
    OrderBuilder = None  # type: ignore[assignment]
    SCHWAB_IMPORT_ERROR = e

# ---------------------------------------------------------------------------
# Paths / config
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
LIVE_INFO_SHEET = "Prices"

DRY_RUN = False  # When True, only print planned exits; do not send orders.


def normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    s = str(direction_cell).strip().lower()
    if s in ("long", "short"):
        return s
    return s


def _tos_exit_cell_to_order_type(cell_value) -> str:
    """
    Column E: ToS Exit.
    'Open' (case-insensitive) -> 'MKT', else -> 'MOC'.
    """
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def read_exit_trade_info(sheet) -> List[Tuple[str, str, int, str]]:
    """
    Read columns A–C and E from sheet (Prices).

    Returns exits = [(ticker, action, size, order_type)], where:
      - action: 'SELL' for long, 'BUY' for short (we map BUY to buy-to-cover)
      - order_type: 'MKT' or 'MOC' based on column E (ToS Exit).
    """
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000

    exits: List[Tuple[str, str, int, str]] = []

    for row in range(2, max_row + 1):
        ticker_cell = sheet.range(f"A{row}").value
        direction_cell = sheet.range(f"B{row}").value
        size_cell = sheet.range(f"C{row}").value
        tos_exit_cell = sheet.range(f"E{row}").value

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

        order_type = _tos_exit_cell_to_order_type(tos_exit_cell)
        action = "SELL" if direction_norm == "long" else "BUY"
        exits.append((ticker, action, size, order_type))

    return exits


def place_exit_orders_schwab(client, account_id: str, exits: List[Tuple[str, str, int, str]]) -> None:
    if SCHWAB_IMPORT_ERROR is not None or OrderType is None or EquityInstruction is None or OrderBuilder is None:
        raise ImportError(
            "Could not import Schwab order classes from schwab-py.\n"
            "Install/update schwab-py with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_IMPORT_ERROR}"
        )

    if not exits:
        print("No exit orders to place.")
        return

    print("\nPlanned Schwab exit orders (close/cover):")
    for ticker, action, size, order_type in exits:
        print(f"  {action} {size} {ticker}  [{order_type}]")

    if DRY_RUN:
        print("\nDRY_RUN is True: no Schwab exit orders will be sent. "
              "Set DRY_RUN = False at the top of this script to send live orders.")
        return

    print("\nPlacing Schwab exit orders...")
    for ticker, action, size, order_type in exits:
        try:
            # Map action string to EquityInstruction
            if action == "SELL":
                instr = EquityInstruction.SELL
            else:  # BUY to close short
                instr = EquityInstruction.BUY_TO_COVER

            # Map our string order_type to Schwab OrderType
            if order_type == "MKT":
                ot = OrderType.MARKET
            else:
                ot = OrderType.MARKET_ON_CLOSE

            ob = OrderBuilder().set_order_type(ot)
            ob = ob.add_equity_leg(instr, ticker, size)
            order_spec = ob.build()
            resp = client.place_order(account_id, order_spec)
            status = resp.status_code if hasattr(resp, "status_code") else resp
            print(f"Submitted {action} {size} {ticker} ({order_type}), response: {status}")
        except Exception as e:
            print(f"Error placing Schwab exit order for {ticker}: {e}")


def main() -> None:
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

        exits = read_exit_trade_info(sheet)

        if not exits:
            print("No valid rows in Live_Trade_Info; nothing to exit.")
            wb.close()
            return

        # Single confirmation prompt before sending
        reply = input("\nSend live exit orders to Schwab? (y/n): ").strip().lower()
        if reply not in ("y", "yes"):
            print("Exiting without sending Schwab exit orders.")
            wb.close()
            return

        try:
            client, cfg = create_client()
        except Exception as e:
            print(f"Failed to create Schwab client: {e}")
            wb.close()
            return

        account_id = cfg.get("account_id")
        if not account_id:
            print("account_id is missing from schwab_config.json; cannot place Schwab exit orders.")
            wb.close()
            return

        place_exit_orders_schwab(client, account_id, exits)

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

