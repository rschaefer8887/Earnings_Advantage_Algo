"""
xlwings helpers for the Latest Earnings workbook.

Attach to an already-open workbook in Excel when possible; otherwise open in a
hidden Excel instance. Scripts that own the instance should call
release_earnings_workbook() in a finally block.
"""

from __future__ import annotations

import os
import time
from typing import Any, Optional, Tuple

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

LATEST_EARNINGS_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Document.xlsx")
LATEST_EARNINGS_SHEET = "Trades"


def _norm_basename(path: str) -> str:
    return os.path.normcase(os.path.basename(path.strip()))


def _file_stem(path: str) -> str:
    return os.path.normcase(os.path.splitext(os.path.basename(path.strip()))[0])


def same_workbook_path(book_path: str, target_path: str) -> bool:
    """Match an xlwings book path to our earnings file path."""
    if not book_path or not target_path:
        return False

    book_str = str(book_path).strip()
    target_abs = os.path.abspath(target_path)

    try:
        book_abs = os.path.abspath(book_str)
    except (OSError, TypeError, ValueError):
        book_abs = book_str

    if os.path.normcase(book_abs) == os.path.normcase(target_abs):
        return True

    try:
        if os.path.exists(book_abs) and os.path.exists(target_abs):
            if os.path.samefile(book_abs, target_abs):
                return True
    except (OSError, ValueError):
        pass

    target_base = _norm_basename(target_abs)
    book_base = _norm_basename(book_str)
    if target_base and book_base == target_base:
        return True

    return False


def _book_matches_earnings(book: Any, abs_path: str) -> bool:
    try:
        if same_workbook_path(book.fullname, abs_path):
            return True
    except Exception:
        pass

    target_stem = _file_stem(abs_path)
    try:
        book_name = str(book.name).strip()
        if book_name and os.path.normcase(book_name) == target_stem:
            return True
    except Exception:
        pass

    return False


def _iter_open_books():
    """Yield xlwings Book objects from every running Excel instance."""
    import xlwings as xw

    try:
        apps = xw.apps
    except Exception:
        return

    for app in apps:
        try:
            for book in app.books:
                yield book
        except Exception:
            continue


def find_open_earnings_workbook(earnings_file: str) -> Optional[Any]:
    """Return the xlwings Book if Latest Earnings is open in any Excel instance."""
    abs_path = os.path.abspath(earnings_file)
    for book in _iter_open_books():
        try:
            if _book_matches_earnings(book, abs_path):
                return book
        except Exception:
            continue
    return None


def _configure_owned_excel_app(app: Any) -> None:
    """Headless Excel settings so save does not block on dialogs."""
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.screen_updating = False
    except Exception:
        pass


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
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"Earnings workbook not found: {abs_path}")

    app = xw.App(visible=False)
    _configure_owned_excel_app(app)

    try:
        wb = app.books.open(abs_path)
    except Exception as open_err:
        try:
            app.quit()
        except Exception:
            pass
        existing = find_open_earnings_workbook(earnings_file)
        if existing is not None:
            print("Using open Latest Earnings Document in Excel.")
            return None, existing, False, False
        raise RuntimeError(
            "Could not open or attach to Latest Earnings workbook. "
            f"Close it in Excel and retry, or check the path: {abs_path}"
        ) from open_err

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


def nudge_onedrive_sync(file_path: str, *, settle_seconds: float = 0.75) -> None:
    """
    Wake OneDrive after a headless Excel COM save.

    Hidden Excel often updates the local file without the sync client noticing
    until something else touches it. Bumping mtime after Excel has released the
    handle (no UI) usually starts upload sooner. Safe no-op if the path is
    missing or still locked.
    """
    abs_path = os.path.abspath(file_path)
    if not os.path.isfile(abs_path):
        return
    if settle_seconds > 0:
        time.sleep(settle_seconds)
    try:
        now = time.time()
        os.utime(abs_path, (now, now))
        print(f"  Nudged OneDrive sync: {os.path.basename(abs_path)}")
    except OSError as e:
        print(f"  Warning: could not nudge OneDrive sync ({e})")
