"""
Shared helper for Live_Trade_Info.xlsx: resolve which sheet to use for trade data.

Reads the "Trade_Mode" sheet, cell C3. If the value is "Single Day", returns
"Daily_Trades"; otherwise returns the current weekday name (Monday, Tuesday, etc.)
so scripts can use Monday/Tuesday/Wednesday/Thursday sheets for multi-day mode.

If the file or sheet is missing, or the cell cannot be read, returns "Daily_Trades".
"""

import os
from datetime import datetime

TRADE_MODE_SHEET = "Trade_Mode"
TRADE_MODE_CELL = "C3"
MODE_SINGLE_DAY = "Single Day"
DEFAULT_SHEET = "Daily_Trades"


def get_trade_sheet_name(file_path: str) -> str:
    """
    Return the Live_Trade_Info sheet name to use for reading or writing trade data.

    - If Trade_Mode!C3 (stripped, case-insensitive) is "Single Day" -> "Daily_Trades".
    - Otherwise -> current weekday name (e.g. "Monday", "Tuesday") for multi-day mode.
    - If file missing, sheet missing, or read error -> "Daily_Trades".
    """
    if not file_path or not os.path.exists(file_path):
        return DEFAULT_SHEET
    try:
        from openpyxl import load_workbook
        wb = load_workbook(file_path, data_only=True)
        if TRADE_MODE_SHEET not in wb.sheetnames:
            return DEFAULT_SHEET
        ws = wb[TRADE_MODE_SHEET]
        value = ws[TRADE_MODE_CELL].value
    except Exception:
        return DEFAULT_SHEET
    if value is None:
        return DEFAULT_SHEET
    if str(value).strip().lower() == MODE_SINGLE_DAY.lower():
        return DEFAULT_SHEET
    return datetime.now().strftime("%A")
