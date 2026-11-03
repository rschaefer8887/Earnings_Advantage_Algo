"""
xlwings helpers for the Latest Earnings workbook.

Attach to an already-open workbook in Excel when possible; otherwise open in a
hidden Excel instance. Scripts that own the instance should call
release_earnings_workbook() in a finally block.

Prefer: open → read → release (if owned) → long API work → open → write/save →
release. Holding Excel open across Schwab fetches causes OneDrive
"someone else has it open" popups.
"""

from __future__ import annotations

import os
import shutil
import time
from datetime import datetime
from typing import Any, Optional, Tuple
from zoneinfo import ZoneInfo

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

# Live workbook kept outside OneDrive to avoid sync/merge overwrites.
LATEST_EARNINGS_DIR = r"C:\Users\ryanc\! -- Trading"
LATEST_EARNINGS_FILE = os.path.join(
    LATEST_EARNINGS_DIR, "! -- Latest Earnings Season.xlsx"
)
LATEST_EARNINGS_SHEET = "Trades"
LATEST_EARNINGS_BACKUP_DIR = os.path.join(LATEST_EARNINGS_DIR, "zBackups")

_MOUNTAIN_TZ = ZoneInfo("America/Denver")


def backup_earnings_workbook(
    file_path: Optional[str] = None,
    *,
    backup_dir: Optional[str] = None,
    label: Optional[str] = None,
) -> Optional[str]:
    """
    Copy the live Earnings workbook into zBackups with a Mountain timestamp.
    Optional label (e.g. 'preopens', 'precloses', 'preMKT prices') is appended
    before .xlsx. Returns backup path, or None on failure.
    """
    src = os.path.abspath(file_path or LATEST_EARNINGS_FILE)
    dest_dir = backup_dir or LATEST_EARNINGS_BACKUP_DIR
    if not os.path.isfile(src):
        print(f"  Warning: backup skipped; missing {src}")
        return None
    try:
        os.makedirs(dest_dir, exist_ok=True)
    except OSError as e:
        print(f"  Warning: could not create backup dir ({e})")
        return None

    stamp = datetime.now(_MOUNTAIN_TZ).strftime("%Y%m%d_%H%M%S")
    suffix = f"_{label}" if label else ""
    dest_name = f"! -- Latest Earnings Season_{stamp}{suffix}.xlsx"
    dest = os.path.join(dest_dir, dest_name)
    try:
        shutil.copy2(src, dest)
        print(f"  Backup saved: {dest}")
        return dest
    except OSError as e:
        print(f"  Warning: earnings backup failed ({e})")
        return None


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
        print("Using open Latest Earnings Season in Excel.")
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
            print("Using open Latest Earnings Season in Excel.")
            return None, existing, False, False
        raise RuntimeError(
            "Could not open or attach to Latest Earnings workbook. "
            f"Close it in Excel and retry, or check the path: {abs_path}"
        ) from open_err

    return app, wb, True, True


def office_lock_path(file_path: str) -> str:
    """Office lock companion (~$Name.xlsx) next to the workbook."""
    abs_path = os.path.abspath(file_path)
    directory, name = os.path.split(abs_path)
    return os.path.join(directory, f"~${name}")


def clear_office_lock_file(file_path: str) -> bool:
    """
    Remove a leftover Office ~$ lock file if present.
    Safe no-op when missing or still in use.
    """
    lock = office_lock_path(file_path)
    if not os.path.isfile(lock):
        return False
    try:
        os.remove(lock)
        print(f"  Cleared stale Office lock: {os.path.basename(lock)}")
        return True
    except OSError as e:
        print(f"  Warning: could not remove Office lock ({e})")
        return False


def wait_until_file_unlocked(
    file_path: str, *, timeout_sec: float = 60.0, poll_sec: float = 0.25
) -> bool:
    """
    Return True once the file can be opened for append (Excel released it).
    """
    abs_path = os.path.abspath(file_path)
    if not os.path.isfile(abs_path):
        return True
    deadline = time.time() + max(0.0, timeout_sec)
    while True:
        try:
            with open(abs_path, "a"):
                pass
            return True
        except OSError:
            if time.time() >= deadline:
                return False
            time.sleep(poll_sec)


def release_earnings_workbook(
    app: Optional[Any],
    wb: Optional[Any],
    *,
    owned_app: bool,
    owned_book: bool,
    earnings_file: Optional[str] = None,
) -> None:
    """
    Close workbook and quit Excel only when this script opened them.
    Then wait briefly for the OS handle to drop and clear a stale ~$ lock.
    """
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
        # COM often keeps the file lock for a beat after quit returns.
        time.sleep(0.75)

    if owned_book and earnings_file:
        wait_until_file_unlocked(earnings_file, timeout_sec=20.0)
        clear_office_lock_file(earnings_file)


def nudge_onedrive_sync(file_path: str, *, settle_seconds: float = 1.5) -> None:
    """
    After a headless Excel save/quit: wait until the file is unlocked so OneDrive
    can sync without fighting Excel.

    Intentionally does NOT bump mtime — that used to force sync while the lock
    was still contested and triggered "someone else has it open" popups.
    """
    abs_path = os.path.abspath(file_path)
    if not os.path.isfile(abs_path):
        return
    if settle_seconds > 0:
        time.sleep(settle_seconds)
    if wait_until_file_unlocked(abs_path, timeout_sec=45.0):
        clear_office_lock_file(abs_path)
        print(f"  Earnings file unlocked for sync: {os.path.basename(abs_path)}")
    else:
        print(
            f"  Warning: {os.path.basename(abs_path)} still locked after wait; "
            "OneDrive may show a temporary lock dialog."
        )
