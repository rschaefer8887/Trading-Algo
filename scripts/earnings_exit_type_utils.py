"""
Utilities to re-read exit-type values from the Latest Earnings workbook.

Used by exit scripts so that Live_Trade_Info D/E columns are refreshed from the
Latest Earnings Document just before placing orders.
"""

from __future__ import annotations

import os
from typing import Any, Dict, Optional

from openpyxl import load_workbook


def _base_dir() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


LATEST_EARNINGS_FILE = os.path.join(_base_dir(), "! -- Latest Earnings Document.xlsx")
LATEST_EARNINGS_SHEET = "Trades"

# Stage_Trades_Auto uses:
#   HEADER_ROW = 3 (data starts at row 4)
#   Column O = flag
#   Column A = ticker
#   Column AA = ToS exit
#   Column AB = IBKR exit
HEADER_ROW = 3
COL_FLAG = "O"
COL_TICKER = "A"
COL_TOS_EXIT = "AA"
COL_IBKR_EXIT = "AB"


def _parse_flag(cell_value: Any) -> Optional[str | int]:
    """Return 'T' for T-rows, 1..5 for weekday rows, else None."""
    if cell_value is None:
        return None
    s = str(cell_value).strip().upper()
    if s == "T":
        return "T"
    try:
        n = int(cell_value) if isinstance(cell_value, (int, float)) else int(s)
    except (ValueError, TypeError):
        return None
    if 1 <= n <= 5:
        return n
    return None


def _cell_to_clean_str(cell_value: Any) -> str:
    """Return a trimmed string; blank if None."""
    if cell_value is None:
        return ""
    return str(cell_value).strip()


def read_exit_types_from_latest_earnings(
    *,
    single_day: bool,
    earnings_file: str = LATEST_EARNINGS_FILE,
    earnings_sheet: str = LATEST_EARNINGS_SHEET,
) -> Dict[str, Dict[str, str]]:
    """
    Read exit-type values from Latest Earnings Document.

    Returns a dict keyed by ticker_upper:
        {
          "AAPL": {"tos_exit": "...", "ib_exit": "..."},
          ...
        }

    Filtering:
    - single_day=True  -> include only rows where flag (col O) == "T"
    - single_day=False -> include only rows where flag (col O) is 1..5

    Matching note:
    - If the same ticker appears multiple times in the filtered range, last one wins.
    """
    if not os.path.exists(earnings_file):
        raise FileNotFoundError(f"Latest earnings file not found: {earnings_file}")

    wb = load_workbook(earnings_file, data_only=True, read_only=True)
    try:
        ws = wb[earnings_sheet]
    except KeyError:
        wb.close()
        raise KeyError(f"Worksheet '{earnings_sheet}' not found in {earnings_file}")

    start_row = HEADER_ROW + 1
    max_row = ws.max_row or (HEADER_ROW + 1)

    if max_row < start_row:
        wb.close()
        return {}

    # Batch-read required columns as arrays to avoid per-row cell lookups.
    # This is the main performance improvement over ws[f"{COL}{row}"].value.
    # Each range read returns a tuple-of-tuples with shape (row_count, 1).
    o_vals = [row[0].value for row in ws[f"{COL_FLAG}{start_row}:{COL_FLAG}{max_row}"]]
    a_vals = [row[0].value for row in ws[f"{COL_TICKER}{start_row}:{COL_TICKER}{max_row}"]]
    aa_vals = [row[0].value for row in ws[f"{COL_TOS_EXIT}{start_row}:{COL_TOS_EXIT}{max_row}"]]
    ab_vals = [row[0].value for row in ws[f"{COL_IBKR_EXIT}{start_row}:{COL_IBKR_EXIT}{max_row}"]]

    row_count = len(o_vals)
    lookup: Dict[str, Dict[str, str]] = {}
    for i in range(row_count):
        row_flag_raw = o_vals[i]
        flag_val = _parse_flag(row_flag_raw)
        if flag_val is None:
            continue

        if single_day:
            if flag_val != "T":
                continue
        else:
            if not isinstance(flag_val, int):
                continue

        ticker_val = a_vals[i]
        if ticker_val is None or str(ticker_val).strip() == "":
            continue
        ticker = str(ticker_val).strip().upper()

        lookup[ticker] = {
            "tos_exit": _cell_to_clean_str(aa_vals[i]),
            "ib_exit": _cell_to_clean_str(ab_vals[i]),
        }

    wb.close()
    return lookup

