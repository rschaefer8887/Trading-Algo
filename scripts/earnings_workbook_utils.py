"""
xlwings helpers for the Latest Earnings workbook.

Attach to an already-open workbook in Excel when possible; otherwise open in a
hidden Excel instance. Scripts that own the instance should call
release_earnings_workbook() in a finally block.
"""

from __future__ import annotations

import os
from typing import Any, Optional, Tuple

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

LATEST_EARNINGS_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Document.xlsx")
LATEST_EARNINGS_SHEET = "Trades"


def same_workbook_path(book_path: str, target_path: str) -> bool:
    if not book_path:
        return False
    return os.path.normcase(os.path.abspath(book_path)) == os.path.normcase(
        os.path.abspath(target_path)
    )


def find_open_earnings_workbook(earnings_file: str) -> Optional[Any]:
    """Return the xlwings Book if Latest Earnings is open in any Excel instance."""
    import xlwings as xw

    abs_path = os.path.abspath(earnings_file)
    for book in xw.books:
        try:
            if same_workbook_path(book.fullname, abs_path):
                return book
        except Exception:
            continue
    return None


def open_or_attach_earnings_workbook(
    earnings_file: str,
) -> Tuple[Optional[Any], Any, bool, bool]:
    """
    Open Latest Earnings for read/write via xlwings.

    Returns:
        (app, wb, owned_app, owned_book)
        - Attached to user's open workbook: (None, book, False, False)
        - Opened by this script: (app, book, True, True)
    """
    import xlwings as xw

    existing = find_open_earnings_workbook(earnings_file)
    if existing is not None:
        print("Using open Latest Earnings Document in Excel.")
        return None, existing, False, False

    abs_path = os.path.abspath(earnings_file)
    app = xw.App(visible=False)
    wb = app.books.open(abs_path)
    return app, wb, True, True


def release_earnings_workbook(
    app: Optional[Any],
    wb: Optional[Any],
    *,
    owned_app: bool,
    owned_book: bool,
) -> None:
    """Close workbook and quit Excel only when this script opened them."""
    if owned_book and wb is not None:
        try:
            wb.close()
        except Exception:
            pass
    if owned_app and app is not None:
        try:
            app.quit()
        except Exception:
            pass