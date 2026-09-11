"""Get_Opens_ToS.py - Opening prices from Schwab/ToS API into Latest Earnings.

Standalone script that fetches daily opens via Schwab/ToS API.

If Latest Earnings is already open in Excel, attaches to that workbook; otherwise
opens it in a hidden Excel instance.

Workbook logic (Trades sheet):
  - Column K: earnings date
  - Column P: first "O" starts the range, "0" stops.
  - Column A: ticker per row.
  - Column T: opening price target.
"""

import os
import sys
import time
from typing import Dict, List, Optional, Tuple

try:
    import xlwings as xw
except ImportError:
    xw = None

from Schwab_Auth import create_client
from earnings_workbook_utils import open_or_attach_earnings_workbook, release_earnings_workbook
from price_run_timestamps import write_opens_success

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Document.xlsx")
SOURCE_SHEET = "Trades"
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


def main() -> None:
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return
    if not os.path.exists(SOURCE_FILE):
        print(f"Source file not found: {SOURCE_FILE}")
        return

    app = None
    wb = None
    owned_app = False
    owned_book = False
    try:
        app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(SOURCE_FILE)
        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"Sheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
            return

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
        if count_o != 1 or count_zero != 1:
            print("One clean range is not selected, please clean up your open range and try again.")
            sys.exit(0)

        first_o_row = None
        for row in range(start_row, max_row + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_start_flag(cell_val):
                first_o_row = row
                break

        if first_o_row is None:
            print("No 'O' flag found in column P. Nothing to process.")
            sys.exit(0)

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
            print("No tickers found in rows between 'O' and '0' in column P.")
            sys.exit(0)

        tickers_to_fetch = [t for _, t in to_process]
        print(f"Fetching opening prices for {len(tickers_to_fetch)} ticker(s) from Schwab daily candles...")
        prices = _fetch_opens_via_schwab(tickers_to_fetch)

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
        else:
            wb.save()
            print("Opening prices written to Latest Earnings (saved via Excel).")
            write_opens_success()

        print(
            f"\nTickers resolved for column {COL_OPENING_PRICE}: "
            f"{', '.join(written) if written else '(none)'}"
        )
    finally:
        release_earnings_workbook(app, wb, owned_app=owned_app, owned_book=owned_book)


if __name__ == "__main__":
    main()
