"""
Open_Close_Price_Repairs — One-off repair for M2/M1/C closes and O opens.

Interactive. For each range, enter a YYYY-MM-DD as-of date, or NA to skip:

  M2  -> column V closes for that date (Q flag chain, same rules as Get_Closes)
  M1  -> column U closes
  C   -> column S closes
  O   -> column T opens for that date (P O...0 range, same as Get_Opens)

Does not stamp get_opens/get_closes success files (repair only).

Usage:
  python scripts/Open_Close_Price_Repairs.py
"""

from __future__ import annotations

import os
import time
from datetime import date, datetime, time as dt_time, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple
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
from run_logging import install_run_log

MOUNTAIN_TZ = ZoneInfo("America/Denver")
# Equity session / daily bar labels — use Eastern, not Mountain (UTC midnights
# shift a day earlier in MT and can match "today's" bar to yesterday's as-of).
MARKET_TZ = ZoneInfo("America/New_York")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Season.xlsx")
SOURCE_SHEET = "Trades"
HEADER_ROW = 3
CHECK_END_ROW = 550

COL_TICKER = "A"
COL_Q = "Q"
COL_P = "P"
COL_V = "V"
COL_U = "U"
COL_S = "S"
COL_T = "T"

DRY_RUN = False

MAX_PRICE_HISTORY_PER_MINUTE = 75
_PRICE_HISTORY_PAUSE_SEC = 60.0 / MAX_PRICE_HISTORY_PER_MINUTE


def _save_stamp_mt() -> str:
    now = datetime.now(MOUNTAIN_TZ)
    hour = int(now.strftime("%I"))
    return now.strftime(f"%Y-%m-%d {hour}:%M:%S %p MT")


def _normalize_ticker(cell_value: Any) -> Optional[str]:
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _is_m2(cell_value: Any) -> bool:
    return cell_value is not None and str(cell_value).strip().upper() == "M2"


def _is_m1(cell_value: Any) -> bool:
    return cell_value is not None and str(cell_value).strip().upper() == "M1"


def _is_c(cell_value: Any) -> bool:
    return cell_value is not None and str(cell_value).strip().upper() == "C"


def _is_q_stop(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    if isinstance(cell_value, (int, float)):
        return cell_value == 0
    return str(cell_value).strip() == "0"


def _is_o_start(cell_value: Any) -> bool:
    return cell_value is not None and str(cell_value).strip().upper() == "O"


def _is_p_stop(cell_value: Any) -> bool:
    return _is_q_stop(cell_value)


def _prompt_as_of(label: str) -> Optional[date]:
    """
    Prompt for YYYY-MM-DD or NA. Returns date or None to skip.
    """
    while True:
        raw = input(
            f"{label} as-of date (YYYY-MM-DD), or NA to skip: "
        ).strip()
        if raw.upper() == "NA" or raw == "":
            if raw == "":
                print("  (blank treated as NA — skipping)")
            return None
        try:
            return date.fromisoformat(raw)
        except ValueError:
            print("  Invalid date. Use YYYY-MM-DD or NA.")


def _candle_session_date(candle: Dict[str, Any]) -> Optional[date]:
    """Trading-session calendar date for a daily candle (America/New_York)."""
    raw = candle.get("datetime")
    if raw is None:
        return None
    try:
        if isinstance(raw, (int, float)):
            ms = float(raw)
            if ms > 1e12:
                ms /= 1000.0
            dt = datetime.fromtimestamp(ms, tz=timezone.utc).astimezone(MARKET_TZ)
            return dt.date()
        s = str(raw).replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(MARKET_TZ).date()
    except (TypeError, ValueError, OSError):
        return None


def _extract_field_for_date(
    history_payload: Any, as_of: date, field: str
) -> Tuple[Optional[float], Optional[date]]:
    """
    Return (price, session_date) for the candle whose ET session date == as_of.
    """
    if not isinstance(history_payload, dict):
        return None, None
    candles = history_payload.get("candles")
    if not isinstance(candles, list) or not candles:
        return None, None
    for candle in candles:
        if not isinstance(candle, dict):
            continue
        session = _candle_session_date(candle)
        if session != as_of:
            continue
        raw = candle.get(field)
        if raw is None:
            continue
        try:
            return round(float(raw), 2), session
        except (TypeError, ValueError):
            continue
    return None, None


def _fetch_daily_history(client: Any, ticker: str, as_of: date) -> Any:
    start = datetime.combine(as_of - timedelta(days=3), dt_time(0, 0), tzinfo=MARKET_TZ)
    end = datetime.combine(as_of + timedelta(days=2), dt_time(23, 59, 59), tzinfo=MARKET_TZ)
    try:
        resp = client.get_price_history_every_day(
            ticker,
            start_datetime=start,
            end_datetime=end,
            need_extended_hours_data=False,
            need_previous_close=False,
        )
        return resp.json() if hasattr(resp, "json") else {}
    except TypeError:
        try:
            resp = client.get_price_history_every_day(ticker)
            return resp.json() if hasattr(resp, "json") else {}
        except Exception:
            return {}
    except Exception:
        return {}


def _fetch_prices_for_date(
    client: Any,
    tickers: List[str],
    as_of: date,
    field: str,
) -> Dict[str, Optional[float]]:
    unique = list(dict.fromkeys(tickers))
    result: Dict[str, Optional[float]] = {}
    n = len(unique)
    for i, ticker in enumerate(unique):
        data = _fetch_daily_history(client, ticker, as_of)
        price, session = _extract_field_for_date(data, as_of, field)
        result[ticker] = price
        if price is None:
            # Help debug mismatches: show ET session dates present in the payload.
            dates: List[str] = []
            if isinstance(data, dict):
                for c in data.get("candles") or []:
                    if isinstance(c, dict):
                        d = _candle_session_date(c)
                        if d is not None:
                            dates.append(d.isoformat())
            hint = f" (candles ET: {', '.join(dates)})" if dates else ""
            print(f"  Warning: no {field} for {ticker} on {as_of}{hint}")
        else:
            print(f"  {ticker}: {field}={price} session={session}")
        if i + 1 < n:
            time.sleep(_PRICE_HISTORY_PAUSE_SEC)
    return result


def _collect_close_jobs(
    sheet: Any, max_row: int
) -> List[Tuple[int, str, str, str]]:
    """
    Same Q walk as Get_Closes: M2->V, M1->U, C->S until Q 0.
    Returns (row, ticker, dest_col, kind) kind in m2|m1|c.
    """
    state: Optional[str] = None  # v|u|s
    kind: Optional[str] = None
    jobs: List[Tuple[int, str, str, str]] = []
    for row in range(HEADER_ROW + 1, max_row + 1):
        flag = sheet.range(f"{COL_Q}{row}").value
        if _is_q_stop(flag):
            break
        if _is_m2(flag):
            state, kind = "v", "m2"
        elif _is_m1(flag):
            state, kind = "u", "m1"
        elif _is_c(flag):
            state, kind = "s", "c"
        if state is None or kind is None:
            continue
        ticker = _normalize_ticker(sheet.range(f"{COL_TICKER}{row}").value)
        if not ticker:
            continue
        dest = COL_V if state == "v" else COL_U if state == "u" else COL_S
        jobs.append((row, ticker, dest, kind))
    return jobs


def _collect_open_jobs(sheet: Any, max_row: int) -> List[Tuple[int, str]]:
    """Same P O...0 walk as Get_Opens. Returns (row, ticker)."""
    count_o = 0
    count_zero = 0
    end_check = min(max_row, CHECK_END_ROW)
    start_row = HEADER_ROW + 1
    for row in range(start_row, end_check + 1):
        cell = sheet.range(f"{COL_P}{row}").value
        if _is_o_start(cell):
            count_o += 1
        elif _is_p_stop(cell):
            count_zero += 1

    first_o = None
    for row in range(start_row, max_row + 1):
        if _is_o_start(sheet.range(f"{COL_P}{row}").value):
            first_o = row
            break

    if first_o is None:
        print("  O range: no P 'O' flag found.")
        return []
    if count_o != 1 or count_zero != 1:
        print(
            f"  O range: unclean P range (O={count_o}, 0={count_zero}); skip opens."
        )
        return []

    jobs: List[Tuple[int, str]] = []
    for row in range(first_o, max_row + 1):
        if _is_p_stop(sheet.range(f"{COL_P}{row}").value):
            break
        ticker = _normalize_ticker(sheet.range(f"{COL_TICKER}{row}").value)
        if ticker:
            jobs.append((row, ticker))
    print(f"  O range: using P O@{first_o} ({len(jobs)} ticker row(s)).")
    return jobs


def main() -> int:
    install_run_log("Open_Close_Price_Repairs")
    print("Open_Close_Price_Repairs — one-off M2/M1/C closes and O opens\n")
    print("Enter YYYY-MM-DD for each range, or NA to skip.\n")

    m2_date = _prompt_as_of("M2 (closes -> V)")
    m1_date = _prompt_as_of("M1 (closes -> U)")
    c_date = _prompt_as_of("C  (closes -> S)")
    o_date = _prompt_as_of("O  (opens  -> T)")

    enabled_closes = {
        k: d
        for k, d in (("m2", m2_date), ("m1", m1_date), ("c", c_date))
        if d is not None
    }
    if not enabled_closes and o_date is None:
        print("Nothing selected (all NA). Exiting.")
        return 0

    print("\nPlan:")
    for kind, d in (("M2", m2_date), ("M1", m1_date), ("C", c_date), ("O", o_date)):
        print(f"  {kind}: {d.isoformat() if d else 'skip (NA)'}")

    if xw is None:
        print("ERROR: xlwings not installed.")
        return 1
    if not os.path.exists(SOURCE_FILE):
        print(f"ERROR: missing {SOURCE_FILE}")
        return 1

    try:
        client, _ = create_client()
    except Exception as e:
        print(f"FAIL: Schwab client: {e}")
        return 1

    app = None
    wb = None
    owned_app = False
    owned_book = False
    did_save = False
    held_through_fetch = False
    # Pending writes collected while Excel is closed during fetch.
    close_writes: List[Tuple[int, str, str, float, date]] = []
    open_writes: List[Tuple[int, str, float, date]] = []
    try:
        try:
            app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                SOURCE_FILE
            )
        except Exception as e:
            print(f"FAIL: open/attach Latest Earnings: {e}")
            return 1

        if owned_book:
            print("Opened Latest Earnings in hidden Excel (read jobs, then release).")
        else:
            print("Attached to open Latest Earnings.")
            held_through_fetch = True

        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"ERROR: sheet '{SOURCE_SHEET}' not found.")
            return 1

        try:
            max_row = int(sheet.used_range.last_cell.row)
        except Exception:
            max_row = 2000

        selected_closes: List[Tuple[int, str, str, str]] = []
        open_jobs: List[Tuple[int, str]] = []

        if enabled_closes:
            close_jobs = _collect_close_jobs(sheet, max_row)
            selected_closes = [j for j in close_jobs if j[3] in enabled_closes]
            print(
                f"\nClose jobs: {len(close_jobs)} in Q chain; "
                f"{len(selected_closes)} selected for repair."
            )

        if o_date is not None:
            open_jobs = _collect_open_jobs(sheet, max_row)

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

        if selected_closes:
            by_date: Dict[date, List[Tuple[int, str, str]]] = {}
            for row, ticker, dest, kind in selected_closes:
                as_of = enabled_closes[kind]
                by_date.setdefault(as_of, []).append((row, ticker, dest))

            for as_of, jobs in by_date.items():
                tickers = [t for _, t, _ in jobs]
                print(
                    f"\nFetching closes for {len(tickers)} ticker(s) as-of {as_of}..."
                )
                prices = _fetch_prices_for_date(client, tickers, as_of, "close")
                for row, ticker, dest in jobs:
                    price = prices.get(ticker)
                    if price is None:
                        continue
                    print(f"  Row {row} {ticker}: {dest} <- {price} (close {as_of})")
                    close_writes.append((row, ticker, dest, price, as_of))

        if o_date is not None and open_jobs:
            tickers = [t for _, t in open_jobs]
            print(
                f"\nFetching opens for {len(tickers)} ticker(s) as-of {o_date}..."
            )
            prices = _fetch_prices_for_date(client, tickers, o_date, "open")
            for row, ticker in open_jobs:
                price = prices.get(ticker)
                if price is None:
                    continue
                print(f"  Row {row} {ticker}: T <- {price} (open {o_date})")
                open_writes.append((row, ticker, price, o_date))

        written = len(close_writes) + len(open_writes)
        if DRY_RUN:
            print(f"\nDRY_RUN: would write {written} cell(s); no save.")
            return 0

        if written == 0:
            print("FAIL: nothing written. Not saving.")
            return 1

        if not held_through_fetch:
            try:
                app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                    SOURCE_FILE
                )
            except Exception as e:
                print(f"FAIL: re-open Latest Earnings for write: {e}")
                return 1
            if owned_book:
                print("Re-opened Latest Earnings in hidden Excel for write/save.")
            try:
                sheet = wb.sheets[SOURCE_SHEET]
            except Exception:
                print(f"ERROR: sheet '{SOURCE_SHEET}' not found.")
                return 1
        else:
            sheet = wb.sheets[SOURCE_SHEET]

        for row, ticker, dest, price, as_of in close_writes:
            sheet.range(f"{dest}{row}").value = price
        for row, ticker, price, as_of in open_writes:
            sheet.range(f"{COL_T}{row}").value = price

        try:
            wb.save()
        except Exception as e:
            print(f"FAIL: save error: {e}")
            return 1
        did_save = True
        print(
            f"\nSAVE OK: wrote {written} cell(s). "
            f"Save timestamp: {_save_stamp_mt()}"
        )
        return 0
    except Exception as e:
        print(f"FAIL: Open_Close_Price_Repairs crashed: {e}")
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