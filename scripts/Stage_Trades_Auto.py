"""
Stage Trades Auto — Build Live_Trade_Info from Latest Earnings using column J flags

Reads the Latest Earnings workbook (Trades sheet). Column J contains flags:
  - Rows with "T" (letter T) in column J are included in the trade list.
  - For each such row: ticker from A, direction from Y, share size from Z.

Writes to Live_Trade_Info.xlsx (same layout as Obtain_Live_Trade_Info):
  - Row 1: A1=Ticker, B1=Direction, C1=Share Size, D1=IBKR Exit, E1=ToS Exit
  - Rows 2+: one row per trade; column D is "open", column E is "MOC" (for Think or Swim exit type).
"""

import os
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", message=".*Unknown extension.*", category=UserWarning)
warnings.filterwarnings("ignore", message=".*Conditional Formatting extension.*", category=UserWarning)

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
COL_FLAG = "J"
COL_TICKER = "A"
COL_DIRECTION = "Y"
COL_SIZE = "Z"

OUTPUT_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
OUTPUT_SHEET = "Prices"
HEADER_D1 = "IBKR Exit"
HEADER_E1 = "ToS Exit"
EXIT_TYPE_OPEN = "open"
TOS_EXIT_MOC = "MOC"


def _is_t_flag(cell_value) -> bool:
    """True if cell is the letter T (include this row). Case-insensitive."""
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "T"


def _normalize_ticker(cell_value):
    if cell_value is None:
        return None
    return str(cell_value).strip().upper()


def _normalize_direction(cell_value):
    if cell_value is None:
        return None
    return str(cell_value).strip().lower()


def main():
    if not os.path.exists(SOURCE_FILE):
        print(f"Source Earnings file not found: {SOURCE_FILE}")
        return

    try:
        wb_source = load_workbook(SOURCE_FILE, data_only=True)
    except PermissionError:
        print("Please close Latest Earnings Document.")
        return

    try:
        ws_source = wb_source[SOURCE_SHEET]
    except KeyError:
        print(f"Worksheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
        return

    max_row = ws_source.max_row
    start_row = HEADER_ROW + 1

    # Collect every row where column J has "T"
    trades = []  # list of (ticker, direction, size)
    for row in range(start_row, max_row + 1):
        flag_cell = ws_source[f"{COL_FLAG}{row}"].value
        if not _is_t_flag(flag_cell):
            continue

        raw_ticker = ws_source[f"{COL_TICKER}{row}"].value
        raw_direction = ws_source[f"{COL_DIRECTION}{row}"].value
        raw_size = ws_source[f"{COL_SIZE}{row}"].value

        ticker = _normalize_ticker(raw_ticker)
        direction = _normalize_direction(raw_direction)
        size = raw_size

        if not ticker:
            print(f"Row {row}: missing ticker; skipping.")
            continue
        if not direction:
            print(f"Row {row}: missing trade direction for ticker {ticker}; skipping.")
            continue
        if size is None or str(size).strip() == "":
            print(f"Row {row}: missing share size for ticker {ticker}; skipping.")
            continue

        trades.append((ticker, direction, size))

    if not trades:
        print("No valid trades found (no rows with 'T' in column J had ticker, direction, and share size).")
        return

    print(f"Collected {len(trades)} trade(s) from Earnings (column J = T).")

    # Load or create Live_Trade_Info
    if os.path.exists(OUTPUT_FILE):
        wb_output = load_workbook(OUTPUT_FILE)
        ws_output = wb_output[OUTPUT_SHEET] if OUTPUT_SHEET in wb_output.sheetnames else wb_output.active
    else:
        wb_output = Workbook()
        ws_output = wb_output.active
        ws_output.title = OUTPUT_SHEET

    # Header row: Ticker, Direction, Share Size, IBKR Exit, ToS Exit
    ws_output["A1"] = "Ticker"
    ws_output["B1"] = "Direction"
    ws_output["C1"] = "Share Size"
    ws_output["D1"] = HEADER_D1
    ws_output["E1"] = HEADER_E1

    # Clear existing data rows (2+)
    if ws_output.max_row > 1:
        ws_output.delete_rows(2, ws_output.max_row - 1)

    # Write one row per trade; column D = "open", column E = "MOC"
    for ticker, direction, size in trades:
        next_row = ws_output.max_row + 1
        ws_output.cell(row=next_row, column=1, value=ticker)
        ws_output.cell(row=next_row, column=2, value=direction)
        ws_output.cell(row=next_row, column=3, value=size)
        ws_output.cell(row=next_row, column=4, value=EXIT_TYPE_OPEN)
        ws_output.cell(row=next_row, column=5, value=TOS_EXIT_MOC)

    # Left-align headers and data (A through E)
    left_align = Alignment(horizontal="left")
    for col in range(1, 6):
        ws_output.cell(row=1, column=col).alignment = left_align
    for row in range(2, ws_output.max_row + 1):
        for col in range(1, 6):
            ws_output.cell(row=row, column=col).alignment = left_align

    wb_output.save(OUTPUT_FILE)
    print(f"Wrote {len(trades)} trade(s) to '{OUTPUT_FILE}' (sheet '{ws_output.title}').")
    print(f"Tickers: {', '.join(t[0] for t in trades)}.")


if __name__ == "__main__":
    main()
