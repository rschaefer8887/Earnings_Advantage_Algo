"""
Get_Actual_MKT_Open_Prices - Write actual MKT open-exit fills into column AE.

Standalone (not wired into Get_Opens_ToS). Uses the same column-P range as
Get_Opens_ToS: start at the sole "O", stop before "0". Within that range, for
rows where AB is "Open", fetch FILLED exit orders from ToS2 (account_id2) and
write average fill price to AE.

Direction for exit side comes from column Y (same as ToS2 / Live column B):
  long  -> match SELL / SELL_TO_CLOSE fills
  short -> match BUY / BUY_TO_COVER fills

Usage:
  python scripts/Get_Actual_MKT_Open_Prices.py

Set DRY_RUN = True to list matches without writing/saving.
"""

from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
    import xlwings as xw
except ImportError:
    xw = None

try:
    from schwab.client import Client
except Exception as e:  # pragma: no cover
    Client = None  # type: ignore[assignment, misc]
    SCHWAB_IMPORT_ERROR = e
else:
    SCHWAB_IMPORT_ERROR = None

from Schwab_Auth import create_client_from_token
from earnings_workbook_utils import (
    nudge_onedrive_sync,
    open_or_attach_earnings_workbook,
    release_earnings_workbook,
)
from run_logging import install_run_log

MOUNTAIN_TZ = ZoneInfo("America/Denver")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Season.xlsx")
SOURCE_SHEET = "Trades"
HEADER_ROW = 3

COL_TICKER = "A"
COL_FLAG = "P"  # same O ... 0 range as Get_Opens_ToS
COL_DIRECTION = "Y"  # ToS2 / Live B
COL_EXIT_TYPE = "AB"  # Open -> MKT exit
COL_FILL = "AE"  # actual fill destination
CHECK_END_ROW = 550

# True = print planned AE writes only; do not modify or save the workbook.
DRY_RUN = False

# How far back to query FILLED orders (Schwab allows up to ~60 days).
LOOKBACK_DAYS = 7

# Prefer MARKET fills; if none for a symbol/side, fall back to any filled exit.
PREFER_MARKET_ORDERS = True

EXIT_INSTR_LONG = {"SELL", "SELL_TO_CLOSE"}
EXIT_INSTR_SHORT = {"BUY", "BUY_TO_COVER", "BUY_TO_CLOSE"}


def _save_stamp_mt() -> str:
    now = datetime.now(MOUNTAIN_TZ)
    hour = int(now.strftime("%I"))
    return now.strftime(f"%Y-%m-%d {hour}:%M:%S %p MT")


def _is_start_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "O"


def _is_stop_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    if isinstance(cell_value, (int, float)):
        return cell_value == 0
    return str(cell_value).strip() == "0"


def _normalize_ticker(cell_value: Any) -> Optional[str]:
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _normalize_direction(cell_value: Any) -> Optional[str]:
    if cell_value is None:
        return None
    s = str(cell_value).strip().lower()
    return s if s in ("long", "short") else None


def _is_open_exit(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().lower() == "open"


def _instruction_side(instruction: str) -> Optional[str]:
    """Map Schwab leg instruction to exit side label: SELL or BUY."""
    u = instruction.strip().upper()
    if u in EXIT_INSTR_LONG:
        return "SELL"
    if u in EXIT_INSTR_SHORT:
        return "BUY"
    return None


def _order_is_market(order: Dict[str, Any]) -> bool:
    ot = str(order.get("orderType") or "").strip().upper()
    return ot in ("MARKET", "MKT")


def _fill_price_from_order(order: Dict[str, Any]) -> Optional[float]:
    """Weighted average fill from execution legs; fallback to order-level price."""
    activities = order.get("orderActivityCollection") or []
    weighted: List[Tuple[float, float]] = []
    for act in activities:
        if str(act.get("activityType") or "").upper() != "EXECUTION":
            continue
        for leg in act.get("executionLegs") or []:
            raw_p = (
                leg.get("averagePrice")
                or leg.get("averageFillPrice")
                or leg.get("price")
            )
            raw_q = leg.get("quantity") or leg.get("filledQuantity") or 0
            try:
                p = float(raw_p)
                q = float(raw_q) if raw_q not in (None, "") else 0.0
            except (TypeError, ValueError):
                continue
            if p > 0:
                weighted.append((p, q if q > 0 else 1.0))

    if weighted:
        tot_q = sum(q for _, q in weighted)
        if tot_q <= 0:
            return round(weighted[-1][0], 4)
        return round(sum(p * q for p, q in weighted) / tot_q, 4)

    for key in ("price", "closePrice", "stopPrice"):
        raw = order.get(key)
        if raw is None or raw == "":
            continue
        try:
            p = float(raw)
        except (TypeError, ValueError):
            continue
        if p > 0:
            return round(p, 4)
    return None


def _order_entered_time(order: Dict[str, Any]) -> datetime:
    raw = (
        order.get("enteredTime")
        or order.get("closeTime")
        or order.get("filledTime")
        or ""
    )
    if not raw:
        return datetime.min.replace(tzinfo=timezone.utc)
    try:
        # Schwab often returns e.g. 2026-10-29T13:30:00+0000
        s = str(raw).replace("Z", "+00:00")
        if len(s) >= 5 and (s[-5] in "+-") and s[-3] != ":":
            s = s[:-2] + ":" + s[-2:]
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except ValueError:
        return datetime.min.replace(tzinfo=timezone.utc)


def _iter_order_legs(order: Dict[str, Any]) -> List[Tuple[str, str]]:
    """Return list of (symbol, side SELL|BUY) for equity legs."""
    out: List[Tuple[str, str]] = []
    for leg in order.get("orderLegCollection") or []:
        instr = leg.get("instrument") or {}
        symbol = _normalize_ticker(instr.get("symbol"))
        side = _instruction_side(str(leg.get("instruction") or ""))
        if symbol and side:
            out.append((symbol, side))
    return out


def _build_fill_index(
    orders: List[Dict[str, Any]],
) -> Dict[Tuple[str, str], List[Tuple[datetime, float, bool]]]:
    """
    Map (symbol, SELL|BUY) -> list of (entered_time, fill_price, is_market)
    newest first after sort.
    """
    index: Dict[Tuple[str, str], List[Tuple[datetime, float, bool]]] = {}
    for order in orders:
        if not isinstance(order, dict):
            continue
        status = str(order.get("status") or "").upper()
        if status not in ("FILLED", "FILLED_PARTIAL", ""):
            # Keep FILLED primarily; some payloads omit after filter
            if status and status != "FILLED":
                continue
        fill = _fill_price_from_order(order)
        if fill is None:
            continue
        when = _order_entered_time(order)
        is_mkt = _order_is_market(order)
        for symbol, side in _iter_order_legs(order):
            index.setdefault((symbol, side), []).append((when, fill, is_mkt))

    for key in index:
        index[key].sort(key=lambda t: t[0], reverse=True)
    return index


def _pick_fill(
    index: Dict[Tuple[str, str], List[Tuple[datetime, float, bool]]],
    ticker: str,
    direction: str,
) -> Optional[Tuple[float, str]]:
    """Return (price, note) or None."""
    side = "SELL" if direction == "long" else "BUY"
    candidates = index.get((ticker, side)) or []
    if not candidates:
        return None
    if PREFER_MARKET_ORDERS:
        mkt = [c for c in candidates if c[2]]
        if mkt:
            return mkt[0][1], "MARKET"
    return candidates[0][1], "any-filled"


def _fetch_filled_orders(client: Any, account_hash: str) -> List[Dict[str, Any]]:
    end = datetime.now(timezone.utc)
    start = end - timedelta(days=LOOKBACK_DAYS)
    kwargs = {
        "from_entered_datetime": start,
        "to_entered_datetime": end,
    }
    status_arg = None
    if Client is not None:
        try:
            status_arg = Client.Order.Status.FILLED
        except Exception:
            status_arg = "FILLED"

    try:
        if status_arg is not None:
            resp = client.get_orders_for_account(
                account_hash, status=status_arg, **kwargs
            )
        else:
            resp = client.get_orders_for_account(account_hash, **kwargs)
    except TypeError:
        # Older client without datetime kwargs
        resp = client.get_orders_for_account(account_hash, status=status_arg)

    code = getattr(resp, "status_code", None)
    if code != 200:
        text = getattr(resp, "text", "") or ""
        raise RuntimeError(f"get_orders_for_account failed ({code}): {text[:500]}")

    data = resp.json()
    if not isinstance(data, list):
        raise RuntimeError(f"Unexpected orders payload type: {type(data).__name__}")
    return [o for o in data if isinstance(o, dict)]


def main() -> int:
    install_run_log("Get_Actual_MKT_Open_Prices")
    print(
        "Get_Actual_MKT_Open_Prices - AE fills for AB=Open "
        f"inside Get_Opens P O..0 range "
        f"(lookback {LOOKBACK_DAYS}d, account_id2 / ToS2)"
    )
    if DRY_RUN:
        print("DRY_RUN is True: will not write AE or save the workbook.\n")

    if xw is None:
        print("ERROR: xlwings is not installed. pip install xlwings")
        return 1
    if SCHWAB_IMPORT_ERROR is not None or Client is None:
        print(
            "ERROR: could not import schwab-py. "
            f"Underlying: {SCHWAB_IMPORT_ERROR}"
        )
        return 1
    if not os.path.exists(SOURCE_FILE):
        print(f"ERROR: Source file not found: {SOURCE_FILE}")
        return 1

    try:
        client, cfg = create_client_from_token()
    except Exception as e:
        print(f"FAIL: Schwab client: {e}")
        return 1

    account_hash = cfg.get("account_id2")
    if not account_hash:
        print("FAIL: account_id2 missing from schwab_config.json (ToS2).")
        return 1

    try:
        orders = _fetch_filled_orders(client, account_hash)
    except Exception as e:
        print(f"FAIL: could not fetch filled orders: {e}")
        return 1

    print(f"Fetched {len(orders)} order payload(s) from ToS2.")
    fill_index = _build_fill_index(orders)
    print(f"Indexed {len(fill_index)} symbol/side fill key(s).")

    app = None
    wb = None
    owned_app = False
    owned_book = False
    did_save = False
    try:
        try:
            app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                SOURCE_FILE
            )
        except Exception as e:
            print(f"FAIL: could not open/attach Latest Earnings: {e}")
            return 1

        if owned_book:
            print("Opened Latest Earnings in hidden Excel.")
        else:
            print("Attached to Latest Earnings already open in Excel.")

        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"ERROR: Sheet '{SOURCE_SHEET}' not found.")
            return 1

        try:
            max_row = int(sheet.used_range.last_cell.row)
        except Exception:
            max_row = 2000
        start_row = HEADER_ROW + 1

        # Same P-column O ... 0 gate as Get_Opens_ToS.
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
            if _is_start_flag(sheet.range(f"{COL_FLAG}{row}").value):
                first_o_row = row
                break

        if first_o_row is None:
            print(
                "FAIL: No 'O' flag found in column P. Nothing to process. "
                "Not saving."
            )
            return 1

        if count_o != 1 or count_zero != 1:
            print(
                "FAIL: One clean Get_Opens range is not selected "
                f"(found O={count_o}, 0={count_zero} in column P). "
                "Clean up P and try again. Not saving."
            )
            return 1

        print(f"Using Get_Opens range starting at P O@{first_o_row} until P 0.")

        candidates = 0
        written = 0
        missed = 0
        skipped_dir = 0
        skipped_not_open = 0

        for row in range(first_o_row, max_row + 1):
            if _is_stop_flag(sheet.range(f"{COL_FLAG}{row}").value):
                break

            ticker = _normalize_ticker(sheet.range(f"{COL_TICKER}{row}").value)
            if not ticker:
                continue

            if not _is_open_exit(sheet.range(f"{COL_EXIT_TYPE}{row}").value):
                skipped_not_open += 1
                continue

            direction = _normalize_direction(
                sheet.range(f"{COL_DIRECTION}{row}").value
            )
            if direction is None:
                print(
                    f"  Row {row} {ticker}: AB=Open but bad direction in Y; skip."
                )
                skipped_dir += 1
                continue

            candidates += 1
            picked = _pick_fill(fill_index, ticker, direction)
            if picked is None:
                print(
                    f"  Row {row} {ticker} ({direction}): no matching "
                    f"{'SELL' if direction == 'long' else 'BUY'} fill on ToS2."
                )
                missed += 1
                continue

            price, note = picked
            existing = sheet.range(f"{COL_FILL}{row}").value
            print(
                f"  Row {row} {ticker} ({direction}): AE <- {price} "
                f"[{note}]"
                + (f" (was {existing})" if existing not in (None, "") else "")
            )
            if not DRY_RUN:
                sheet.range(f"{COL_FILL}{row}").value = price
            written += 1

        print(
            f"\nSummary (O..0 range): AB=Open rows={candidates}, "
            f"AE writes={written}, misses={missed}, "
            f"bad direction={skipped_dir}, "
            f"in-range non-Open AB={skipped_not_open}"
        )

        if DRY_RUN:
            print("DRY_RUN complete - workbook unchanged.")
            return 0

        if written == 0:
            print(
                "FAIL: no AE values written (no AB=Open matches with fills). "
                "Not saving."
            )
            return 1

        try:
            wb.save()
        except Exception as e:
            print(f"FAIL: could not save Latest Earnings: {e}")
            return 1
        did_save = True
        print(f"SAVE OK: AE updated for {written} row(s). Save timestamp: {_save_stamp_mt()}")
        return 0
    except Exception as e:
        print(f"FAIL: Get_Actual_MKT_Open_Prices crashed: {e}")
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