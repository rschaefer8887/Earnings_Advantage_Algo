"""
Get Closes ToS - Closing prices from Schwab/ToS API into Latest Earnings.

Standalone script that fetches daily closes via Schwab/ToS API.

If Latest Earnings is already open in Excel, attaches to that workbook; otherwise
opens it in a hidden Excel instance.

Workbook logic (Trades sheet):
  - Column Q flag states: M2 -> write to V, M1 -> write to U, C -> write to S, 0 stops.
  - Column A: ticker per row.
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
from price_run_timestamps import write_closes_success

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Document.xlsx")
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

        state = None  # "v" | "u" | "s"
        to_process: List[Tuple[int, str, str]] = []
        tickers_to_fetch: List[str] = []

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
            tickers_to_fetch.append(ticker)

        if not to_process:
            print("No closing ranges found.")
            sys.exit(0)

        print(f"Fetching closing prices for {len(tickers_to_fetch)} ticker(s) from Schwab daily candles...")
        prices = _fetch_closes_via_schwab(tickers_to_fetch)

        tickers_by_column: Dict[str, List[str]] = {COL_V: [], COL_U: [], COL_S: []}
        for row, ticker, target_col in to_process:
            price = prices.get(ticker)
            if price is None:
                print(f"  Warning: no closing price for {ticker} (row {row})")
                continue
            if not DRY_RUN:
                sheet.range(f"{target_col}{row}").value = price
            tickers_by_column[target_col].append(ticker)

        if DRY_RUN:
            print("DRY_RUN is True: no workbook writes were made.")
        else:
            wb.save()
            print("Closing prices written to Latest Earnings (saved via Excel).")
            write_closes_success()

        print("\nTickers resolved by column:")
        for col in (COL_V, COL_U, COL_S):
            tickers = tickers_by_column[col]
            print(f"  Column {col}: {', '.join(tickers) if tickers else '(none)'}")
    finally:
        release_earnings_workbook(app, wb, owned_app=owned_app, owned_book=owned_book)


if __name__ == "__main__":
    main()
