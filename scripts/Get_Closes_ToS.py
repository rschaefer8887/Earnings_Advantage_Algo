"""
Get Closes ToS - Closing prices from Schwab/ToS API into Latest Earnings.

Standalone script that fetches daily closes via Schwab/ToS API.

If Latest Earnings is already open in Excel, attaches and saves that session;
otherwise opens a hidden Excel instance. Leaves a user-owned workbook open.

Workbook is released during Schwab fetches when this script owns Excel, so
OneDrive is not locked for minutes.

Workbook logic (Trades sheet):
  - Column Q flag states: M2 -> write to V, M1 -> write to U, C -> write to S, 0 stops.
  - Column A: ticker per row.

Success stamp is written only after prices are written and wb.save() succeeds.
Empty/no-work runs exit non-zero so Task Scheduler does not report success.
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
    nudge_onedrive_sync,
    open_or_attach_earnings_workbook,
    release_earnings_workbook,
)
from price_run_timestamps import write_closes_success
from run_logging import install_run_log

MOUNTAIN_TZ = ZoneInfo("America/Denver")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Season.xlsx")
SOURCE_SHEET = "Trades"
HEADER_ROW = 3
COL_TICKER = "A"
COL_FLAG = "Q"
COL_V = "V"
COL_U = "U"
COL_S = "S"

# Manual-test mode: writes/saves workbook output.
DRY_RUN = False

# One Schwab price-history request per ticker; throttle to reduce rate-limit risk.
MAX_PRICE_HISTORY_PER_MINUTE = 75
_PRICE_HISTORY_PAUSE_SEC = 60.0 / MAX_PRICE_HISTORY_PER_MINUTE


def _is_m2(cell_value) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "M2"


def _is_m1(cell_value) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "M1"


def _is_c(cell_value) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "C"


def _is_stop(cell_value) -> bool:
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


def _extract_latest_candle_close(history_payload) -> Optional[float]:
    if not isinstance(history_payload, dict):
        return None

    candles = history_payload.get("candles")
    if not isinstance(candles, list) or not candles:
        return None

    for candle in reversed(candles):
        if not isinstance(candle, dict):
            continue
        close_val = candle.get("close")
        if close_val is None:
            continue
        try:
            return round(float(close_val), 2)
        except (TypeError, ValueError):
            continue

    return None


def _fetch_symbol_daily_close(client, ticker: str) -> Optional[float]:
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
        return _extract_latest_candle_close(data)
    except TypeError:
        # Some client versions accept fewer arguments.
        try:
            resp = client.get_price_history_every_day(ticker)
            data = resp.json() if hasattr(resp, "json") else {}
            return _extract_latest_candle_close(data)
        except Exception:
            return None
    except Exception:
        return None


def _fetch_closes_via_schwab(tickers: List[str]) -> Dict[str, Optional[float]]:
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
        result[ticker] = _fetch_symbol_daily_close(client, ticker)
        if i + 1 < n:
            time.sleep(_PRICE_HISTORY_PAUSE_SEC)

    return result


def _save_stamp_mt() -> str:
    now = datetime.now(MOUNTAIN_TZ)
    hour = int(now.strftime("%I"))
    return now.strftime(f"%Y-%m-%d {hour}:%M:%S %p MT")


def _collect_close_jobs(sheet) -> List[Tuple[int, str, str]]:
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 2000
    start_row = HEADER_ROW + 1

    state = None  # "v" | "u" | "s"
    to_process: List[Tuple[int, str, str]] = []

    for row in range(start_row, max_row + 1):
        flag_cell = sheet.range(f"{COL_FLAG}{row}").value

        if _is_stop(flag_cell):
            break

        if _is_m2(flag_cell):
            state = "v"
        elif _is_m1(flag_cell):
            state = "u"
        elif _is_c(flag_cell):
            state = "s"

        if state is None:
            continue

        ticker = _normalize_ticker(sheet.range(f"{COL_TICKER}{row}").value)
        if not ticker:
            continue

        target_col = COL_V if state == "v" else COL_U if state == "u" else COL_S
        to_process.append((row, ticker, target_col))

    return to_process


def main() -> int:
    install_run_log("Get_Closes_ToS")
    if xw is None:
        print("ERROR: xlwings is not installed. Install it with: pip install xlwings")
        return 1
    if not os.path.exists(SOURCE_FILE):
        print(f"ERROR: Source file not found: {SOURCE_FILE}")
        return 1

    app = None
    wb = None
    owned_app = False
    owned_book = False
    did_save = False
    exit_code = 1
    to_process: List[Tuple[int, str, str]] = []
    # When we own Excel, release between read and fetch. When attached, keep open.
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

        to_process = _collect_close_jobs(sheet)
        if not to_process:
            print(
                "FAIL: No closing ranges found (no M2/M1/C…0 chain with tickers). "
                "Not recording success."
            )
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

        tickers_to_fetch = [t for _, t, _ in to_process]
        print(
            f"Fetching closing prices for {len(tickers_to_fetch)} ticker(s) "
            f"from Schwab daily candles..."
        )
        prices = _fetch_closes_via_schwab(tickers_to_fetch)

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

        tickers_by_column: Dict[str, List[str]] = {COL_V: [], COL_U: [], COL_S: []}
        written_count = 0
        for row, ticker, target_col in to_process:
            price = prices.get(ticker)
            if price is None:
                print(f"  Warning: no closing price for {ticker} (row {row})")
                continue
            if not DRY_RUN:
                sheet.range(f"{target_col}{row}").value = price
            tickers_by_column[target_col].append(ticker)
            written_count += 1

        if DRY_RUN:
            print("DRY_RUN is True: no workbook writes were made.")
            exit_code = 0
        elif written_count == 0:
            print(
                "FAIL: Get_Closes completed with no prices written (fetch miss). "
                "Not saving; not recording success."
            )
            return 1
        else:
            try:
                wb.save()
            except Exception as e:
                print(f"FAIL: could not save Latest Earnings after writing closes: {e}")
                return 1
            did_save = True
            stamp = _save_stamp_mt()
            print(
                f"SAVE OK: Closing prices written to Latest Earnings "
                f"({written_count} ticker(s)). Save timestamp: {stamp}"
            )
            write_closes_success()
            exit_code = 0

        print("\nTickers resolved by column:")
        for col in (COL_V, COL_U, COL_S):
            tickers = tickers_by_column[col]
            print(f"  Column {col}: {', '.join(tickers) if tickers else '(none)'}")
        return exit_code
    except Exception as e:
        print(f"FAIL: Get_Closes_ToS crashed: {e}")
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
