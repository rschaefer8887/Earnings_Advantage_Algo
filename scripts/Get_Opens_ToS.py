"""Get_Opens_ToS.py - Opening prices from Schwab/ToS API into Latest Earnings.

Standalone script that fetches daily opens via Schwab/ToS API.

If Latest Earnings is already open in Excel, attaches and saves that session;
otherwise opens a hidden Excel instance. Leaves a user-owned workbook open.

Workbook is released during Schwab fetches when this script owns Excel, so
OneDrive is not locked for minutes.

Workbook logic (Trades sheet):
  - Column K: earnings date
  - Column P: first "O" starts the range, "0" stops.
  - Column A: ticker per row.
  - Column T: opening price target.

Success stamp is written only after prices are written and wb.save() succeeds.
No O and no 0 in column P (no open range) exits 0 so Task Scheduler is clean;
success is not stamped in that case.
"""

import os
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
    import xlwings as xw
except ImportError:
    xw = None

from Schwab_Auth import create_client
from earnings_workbook_utils import (
    LATEST_EARNINGS_FILE,
    LATEST_EARNINGS_SHEET,
    backup_earnings_workbook,
    nudge_onedrive_sync,
    open_or_attach_earnings_workbook,
    release_earnings_workbook,
)
from price_run_timestamps import write_opens_success
from run_logging import install_run_log

MOUNTAIN_TZ = ZoneInfo("America/Denver")

SOURCE_FILE = LATEST_EARNINGS_FILE
SOURCE_SHEET = LATEST_EARNINGS_SHEET
HEADER_ROW = 3
COL_TICKER = "A"
COL_FLAG = "P"
COL_OPENING_PRICE = "T"
CHECK_END_ROW = 550

DRY_RUN = False

# One Schwab price-history request per ticker; throttle to reduce rate-limit risk.
MAX_PRICE_HISTORY_PER_MINUTE = 75
_PRICE_HISTORY_PAUSE_SEC = 60.0 / MAX_PRICE_HISTORY_PER_MINUTE


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


def _normalize_ticker(cell_value) -> Optional[str]:
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _extract_latest_candle_open(history_payload) -> Optional[float]:
    if not isinstance(history_payload, dict):
        return None

    candles = history_payload.get("candles")
    if not isinstance(candles, list) or not candles:
        return None

    for candle in reversed(candles):
        if not isinstance(candle, dict):
            continue
        open_val = candle.get("open")
        if open_val is None:
            continue
        try:
            return round(float(open_val), 2)
        except (TypeError, ValueError):
            continue

    return None


def _fetch_symbol_daily_open(client, ticker: str) -> Optional[float]:
    if not hasattr(client, "get_price_history_every_day"):
        return None

    # Try with explicit regular-session candles first.
    try:
        resp = client.get_price_history_every_day(
            ticker,
            need_extended_hours_data=False,
            need_previous_close=False,
        )
        data = resp.json() if hasattr(resp, "json") else {}
        return _extract_latest_candle_open(data)
    except TypeError:
        # Some client versions accept fewer arguments.
        try:
            resp = client.get_price_history_every_day(ticker)
            data = resp.json() if hasattr(resp, "json") else {}
            return _extract_latest_candle_open(data)
        except Exception:
            return None
    except Exception:
        return None


def _fetch_opens_via_schwab(tickers: List[str]) -> Dict[str, Optional[float]]:
    unique = list(dict.fromkeys(tickers))
    if not unique:
        return {}

    try:
        client, _ = create_client()
    except Exception as e:
        print(f"Failed to create Schwab client: {e}")
        return {t: None for t in unique}

    result: Dict[str, Optional[float]] = {}
    n = len(unique)
    for i, ticker in enumerate(unique):
        result[ticker] = _fetch_symbol_daily_open(client, ticker)
        if i + 1 < n:
            time.sleep(_PRICE_HISTORY_PAUSE_SEC)

    return result


def _save_stamp_mt() -> str:
    now = datetime.now(MOUNTAIN_TZ)
    hour = int(now.strftime("%I"))
    return now.strftime(f"%Y-%m-%d {hour}:%M:%S %p MT")


# Sentinel from _collect_open_jobs: no P O and no P 0 at all (valid empty day).
_NO_OPEN_RANGE = "NO_OPEN_RANGE"


def _collect_open_jobs(sheet) -> Tuple[Optional[str], List[Tuple[int, str]]]:
    """
    Validate P O...0 and return (status_or_None, jobs).

    status None = ok; _NO_OPEN_RANGE = no range (clean exit);
    other str = FAIL message (exit 1).
    """
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 2000
    start_row = HEADER_ROW + 1

    count_o = 0
    count_zero = 0
    end_check = min(max_row, CHECK_END_ROW)
    for row in range(start_row, end_check + 1):
        cell_val = sheet.range(f"{COL_FLAG}{row}").value
        if _is_start_flag(cell_val):
            count_o += 1
        elif _is_stop_flag(cell_val):
            count_zero += 1

    first_o_row = None
    for row in range(start_row, max_row + 1):
        cell_val = sheet.range(f"{COL_FLAG}{row}").value
        if _is_start_flag(cell_val):
            first_o_row = row
            break

    # No open range at all — normal when there were no live trades today.
    if count_o == 0 and count_zero == 0:
        return _NO_OPEN_RANGE, []

    if first_o_row is None:
        return (
            "FAIL: No 'O' flag found in column P. Nothing to process. "
            "Not recording success.",
            [],
        )

    if count_o != 1 or count_zero != 1:
        return (
            "FAIL: One clean range is not selected "
            f"(found O={count_o}, 0={count_zero}). "
            "Clean up column P and try again. Not recording success.",
            [],
        )

    to_process: List[Tuple[int, str]] = []
    for row in range(first_o_row, max_row + 1):
        cell_val = sheet.range(f"{COL_FLAG}{row}").value
        if _is_stop_flag(cell_val):
            break
        ticker = _normalize_ticker(sheet.range(f"{COL_TICKER}{row}").value)
        if not ticker:
            continue
        to_process.append((row, ticker))

    if not to_process:
        return (
            "FAIL: No tickers found in rows between 'O' and '0' in column P. "
            "Not recording success.",
            [],
        )

    return None, to_process


def main() -> int:
    install_run_log("Get_Opens_ToS")
    if xw is None:
        print("ERROR: xlwings is not installed. Install it with: pip install xlwings")
        return 1
    if not os.path.exists(SOURCE_FILE):
        print(f"ERROR: Source file not found: {SOURCE_FILE}")
        return 1

    backup_earnings_workbook(SOURCE_FILE, label="preopens")

    app = None
    wb = None
    owned_app = False
    owned_book = False
    did_save = False
    exit_code = 1
    to_process: List[Tuple[int, str]] = []
    held_through_fetch = False

    try:
        try:
            app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                SOURCE_FILE
            )
        except Exception as e:
            print(f"FAIL: could not open or attach to Latest Earnings: {e}")
            return 1

        if owned_book:
            print("Opened Latest Earnings in hidden Excel (read jobs, then release).")
        else:
            print(
                "Attached to Latest Earnings already open in Excel "
                "(will save; leave your window open)."
            )
            held_through_fetch = True

        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"ERROR: Sheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
            return 1

        err, to_process = _collect_open_jobs(sheet)
        if err == _NO_OPEN_RANGE:
            print(
                "No open range in column P (no O and no 0). "
                "No live opens to populate today. Exiting cleanly."
            )
            return 0
        if err:
            print(err)
            return 1

        if owned_book:
            release_earnings_workbook(
                app,
                wb,
                owned_app=owned_app,
                owned_book=owned_book,
                earnings_file=SOURCE_FILE,
            )
            app, wb = None, None
            owned_app, owned_book = False, False
            print("Released Excel during Schwab fetch (avoids OneDrive lock).")

        tickers_to_fetch = [t for _, t in to_process]
        print(
            f"Fetching opening prices for {len(tickers_to_fetch)} ticker(s) "
            f"from Schwab daily candles..."
        )
        prices = _fetch_opens_via_schwab(tickers_to_fetch)

        if not held_through_fetch:
            try:
                app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                    SOURCE_FILE
                )
            except Exception as e:
                print(f"FAIL: could not re-open Latest Earnings for write: {e}")
                return 1
            if owned_book:
                print("Re-opened Latest Earnings in hidden Excel for write/save.")
            try:
                sheet = wb.sheets[SOURCE_SHEET]
            except Exception:
                print(f"ERROR: Sheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
                return 1
        else:
            sheet = wb.sheets[SOURCE_SHEET]

        written: List[str] = []
        for row, ticker in to_process:
            price = prices.get(ticker)
            if price is None:
                print(f"  Warning: no opening price for {ticker} (row {row})")
                continue
            if not DRY_RUN:
                sheet.range(f"{COL_OPENING_PRICE}{row}").value = price
            written.append(ticker)

        if DRY_RUN:
            print("DRY_RUN is True: no workbook writes were made.")
            exit_code = 0
        elif not written:
            print(
                "FAIL: Get_Opens completed with no prices written (fetch miss). "
                "Not saving; not recording success."
            )
            return 1
        else:
            try:
                wb.save()
            except Exception as e:
                print(f"FAIL: could not save Latest Earnings after writing opens: {e}")
                return 1
            did_save = True
            stamp = _save_stamp_mt()
            print(
                f"SAVE OK: Opening prices written to Latest Earnings "
                f"({len(written)} ticker(s)). Save timestamp: {stamp}"
            )
            write_opens_success()
            exit_code = 0

        print(
            f"\nTickers resolved for column {COL_OPENING_PRICE}: "
            f"{', '.join(written) if written else '(none)'}"
        )
        return exit_code
    except Exception as e:
        print(f"FAIL: Get_Opens_ToS crashed: {e}")
        return 1
    finally:
        release_earnings_workbook(
            app,
            wb,
            owned_app=owned_app,
            owned_book=owned_book,
            earnings_file=SOURCE_FILE,
        )
        if did_save:
            nudge_onedrive_sync(SOURCE_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
