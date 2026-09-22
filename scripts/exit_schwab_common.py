"""
Shared Schwab exit helpers for Exit_ToS / Exit_ToS2.

Reads staged exit types from Live_Trade_Info and places close/cover orders.
Does not touch Latest Earnings.
"""

from __future__ import annotations

import asyncio
import os
from datetime import datetime
from typing import List, Optional, Tuple

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

try:
    import xlwings as xw
except ImportError:
    xw = None

from live_trade_info_utils import (
    MODE_FULL_AUTO,
    MODE_SINGLE_DAY,
    TRADE_MODE_CELL,
    TRADE_MODE_SHEET,
)
from full_auto_utils import full_auto_sheet_name_for_today
from Schwab_Auth import create_client

SCHWAB_IMPORT_ERROR = None
try:
    from schwab.orders.common import (
        OrderType,
        EquityInstruction,
        Duration,
        Session,
        OrderStrategyType,
    )
    from schwab.orders.generic import OrderBuilder
except Exception as e:  # pragma: no cover
    OrderType = None  # type: ignore[assignment]
    EquityInstruction = None  # type: ignore[assignment]
    Duration = None  # type: ignore[assignment]
    Session = None  # type: ignore[assignment]
    OrderStrategyType = None  # type: ignore[assignment]
    OrderBuilder = None  # type: ignore[assignment]
    SCHWAB_IMPORT_ERROR = e

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")

DRY_RUN = False


def _normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    return str(direction_cell).strip().lower()


def _exit_cell_to_order_type(cell_value) -> str:
    """'Open' -> MKT; anything else (incl. blank/MOC) -> MOC."""
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def _last_row_by_ticker(sheet, col_letter: str = "A", min_row: int = 2, max_scan: int = 5000) -> int:
    """Last non-empty ticker row; caps scan to used_range to avoid huge COM pulls."""
    try:
        used_last = int(sheet.used_range.last_cell.row)
        max_scan = min(max_scan, max(used_last + 20, min_row + 50))
    except Exception:
        pass
    try:
        vals = sheet.range(f"{col_letter}{min_row}:{col_letter}{max_scan}").value
        last = min_row - 1
        if vals is None:
            return last
        for i, v in enumerate(vals):
            cell_val = v[0] if isinstance(v, list) else v
            if cell_val is None:
                continue
            if str(cell_val).strip() != "":
                last = min_row + i
        return last
    except Exception:
        try:
            return sheet.used_range.last_cell.row
        except Exception:
            return min_row - 1


def read_exit_trade_info(sheet, exit_col: str) -> List[Tuple[str, str, int, str]]:
    """
    Bulk-read A–E once; build exits from A/B/C and exit_col ('D' or 'E').

    Returns [(ticker, action, size, order_type)] with action SELL|BUY and
    order_type MKT|MOC.
    """
    exit_col = exit_col.upper()
    if exit_col not in ("D", "E"):
        raise ValueError(f"exit_col must be 'D' or 'E', got {exit_col!r}")
    exit_idx = 3 if exit_col == "D" else 4  # 0-based within A:E

    max_row = _last_row_by_ticker(sheet, col_letter="A", min_row=2, max_scan=5000)
    if max_row < 2:
        return []

    raw = sheet.range(f"A2:E{max_row}").value
    if raw is None:
        return []
    if not isinstance(raw, list):
        raw = [raw]
    # Single-row range comes back as a flat list of cells, not a list of rows.
    if raw and not isinstance(raw[0], (list, tuple)):
        raw = [raw]

    exits: List[Tuple[str, str, int, str]] = []
    for i, row_vals in enumerate(raw):
        row = 2 + i
        if not isinstance(row_vals, (list, tuple)):
            row_vals = [row_vals]
        # Pad to 5 columns if Excel returns a short row.
        cells = list(row_vals) + [None] * (5 - len(row_vals))
        ticker_cell, direction_cell, size_cell = cells[0], cells[1], cells[2]
        exit_type_cell = cells[exit_idx]

        if ticker_cell is None or str(ticker_cell).strip() == "":
            continue

        ticker = str(ticker_cell).strip().upper()
        direction_norm = _normalize_direction(direction_cell)
        if direction_norm not in ("long", "short"):
            print(f"Row {row}: invalid direction '{direction_cell}' for ticker {ticker}; skipping.")
            continue

        try:
            size = int(size_cell)
        except (TypeError, ValueError):
            print(f"Row {row}: invalid share size '{size_cell}' for ticker {ticker}; skipping.")
            continue

        if size <= 0:
            print(f"Row {row}: non-positive share size {size} for ticker {ticker}; skipping.")
            continue

        order_type = _exit_cell_to_order_type(exit_type_cell)
        action = "SELL" if direction_norm == "long" else "BUY"
        exits.append((ticker, action, size, order_type))

    return exits


def place_exit_orders_schwab(client, account_id: str, exits: List[Tuple[str, str, int, str]]) -> None:
    if (
        SCHWAB_IMPORT_ERROR is not None
        or OrderType is None
        or EquityInstruction is None
        or Duration is None
        or Session is None
        or OrderStrategyType is None
        or OrderBuilder is None
    ):
        raise ImportError(
            "Could not import Schwab order classes from schwab-py.\n"
            "Install/update schwab-py with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_IMPORT_ERROR}"
        )

    if not exits:
        print("No exit orders to place.")
        return

    print("\nPlanned Schwab exit orders (close/cover):")
    for ticker, action, size, order_type in exits:
        print(f"  {action} {size} {ticker}  [{order_type}]")

    if DRY_RUN:
        print(
            "\nDRY_RUN is True: no Schwab exit orders will be sent. "
            "Set DRY_RUN = False at the top of Exit_ToS.py / Exit_ToS2.py to send live orders."
        )
        return

    print("\nPlacing Schwab exit orders...")
    for ticker, action, size, order_type in exits:
        try:
            instr = EquityInstruction.SELL if action == "SELL" else EquityInstruction.BUY_TO_COVER
            ot = OrderType.MARKET if order_type == "MKT" else OrderType.MARKET_ON_CLOSE
            order_spec = (
                OrderBuilder()
                .set_order_type(ot)
                .set_duration(Duration.DAY)
                .set_session(Session.NORMAL)
                .set_order_strategy_type(OrderStrategyType.SINGLE)
                .add_equity_leg(instr, ticker, size)
                .build()
            )
            resp = client.place_order(account_id, order_spec)
            status = getattr(resp, "status_code", None)
            text = getattr(resp, "text", None)
            print(f"Submitted {action} {size} {ticker} ({order_type}), response: {status}")
            print(status, text)
        except Exception as e:
            print(f"Error placing Schwab exit order for {ticker}: {e}")


def _resolve_sheet_name(wb) -> Tuple[Optional[str], bool]:
    """Return (sheet_name, full_auto). sheet_name None means nothing to exit."""
    try:
        trade_mode_value = wb.sheets[TRADE_MODE_SHEET].range(TRADE_MODE_CELL).value
    except Exception:
        trade_mode_value = None

    mode_lower = str(trade_mode_value).strip().lower() if trade_mode_value is not None else ""
    full_auto = mode_lower == MODE_FULL_AUTO.lower()

    if mode_lower == MODE_SINGLE_DAY.lower():
        return "Daily_Trades", full_auto
    if full_auto:
        sheet_name = full_auto_sheet_name_for_today()
        if not sheet_name:
            print("Full_Auto: today is not a trading day; nothing to exit.")
            return None, full_auto
        print(f"Full_Auto: using Live_Trade_Info sheet '{sheet_name}'.")
        return sheet_name, full_auto
    # Emergency / weekday sheet; edit this line to target a past day.
    return datetime.now().strftime("%A"), full_auto

# run_exit_main is the main function that is called from Exit_ToS.py and Exit_ToS2.py
def run_exit_main(*, account_key: str, exit_col: str) -> int:
    """
    Standalone exit run for one Schwab account.

    account_key: 'account_id' (Exit_ToS) or 'account_id2' (Exit_ToS2)
    exit_col: 'E' (primary ToS Exit) or 'D' (secondary / IBKR Exit staging)
    """
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return 1
    if not os.path.exists(LIVE_INFO_FILE):
        print(f"Live trade info file not found: {LIVE_INFO_FILE}")
        return 1

    app = None
    try:
        app = xw.App(visible=False)
        try:
            app.display_alerts = False
        except Exception:
            pass

        try:
            wb = app.books.open(os.path.abspath(LIVE_INFO_FILE))
        except Exception:
            print("Please close Live_Trade_Info")
            return 1

        sheet_name, full_auto = _resolve_sheet_name(wb)
        if not sheet_name:
            wb.close()
            return 0

        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            print(f"Sheet '{sheet_name}' not found in {LIVE_INFO_FILE}.")
            wb.close()
            return 0 if full_auto else 1

        exits = read_exit_trade_info(sheet, exit_col=exit_col)
        # Read-only: close without save.
        wb.close()

        if not exits:
            print(f"No valid exit rows on sheet '{sheet_name}'; nothing to exit.")
            return 0

        if DRY_RUN:
            place_exit_orders_schwab(None, "", exits)
            return 0

        try:
            client, cfg = create_client()
        except Exception as e:
            print(f"Failed to create Schwab client: {e}")
            return 1

        account_id = cfg.get(account_key)
        if not account_id:
            print(
                f"{account_key} is missing from schwab_config.json; "
                "cannot place Schwab exit orders."
            )
            return 1

        place_exit_orders_schwab(client, account_id, exits)
        return 0
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
