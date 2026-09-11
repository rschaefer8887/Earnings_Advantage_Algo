"""
Exit_ToS — Exit live trades via Schwab (ToS) based on Live_Trade_Info.xlsx

Logic mirrors Exit_IB_via_GW, but:
- Sends exit orders to Schwab instead of IB Gateway.
- Uses column E ("ToS Exit") for Schwab exit type:
    - "Open" → MARKET (execute during the session)
    - Anything else (including "MOC" from Stage_Trades_Auto) → MARKET_ON_CLOSE

Workbook shape (sheet "Daily_Trades" in Live_Trade_Info.xlsx):
- Column A: Ticker
- Column B: Direction ("long" / "short")
- Column C: Share Size
- Column D: IB Exit (used by IB exit scripts, ignored here)
- Column E: ToS Exit (used by this script only)

Flow:
- Reads Live_Trade_Info.xlsx via xlwings.
- For each valid row:
    - LONG  -> action SELL   (close long)
    - SHORT -> action BUY_TO_COVER (close short)
- Order type per row:
    - E == "Open" → Schwab OrderType.MARKET
    - else        → Schwab OrderType.MARKET_ON_CLOSE

Prerequisites:
- schwab-py and xlwings installed:
    python -m pip install --upgrade schwab-py xlwings
- Schwab_Auth.py configured and run at least once to complete OAuth.
- Close Live_Trade_Info.xlsx in Excel before running.
"""

import os
import warnings
from typing import List, Tuple

import asyncio

warnings.filterwarnings("ignore", message=".*Unknown extension.*", category=UserWarning)
warnings.filterwarnings(
    "ignore", message=".*Conditional Formatting extension.*", category=UserWarning
)

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

try:
    import xlwings as xw
except ImportError:
    xw = None

from datetime import datetime

from live_trade_info_utils import (
    MODE_FULL_AUTO,
    MODE_SINGLE_DAY,
    TRADE_MODE_CELL,
    TRADE_MODE_SHEET,
    get_trade_sheet_name,
)

from Schwab_Auth import create_client
from earnings_exit_type_utils import read_exit_types_from_latest_earnings

SCHWAB_IMPORT_ERROR = None
try:
    from schwab.orders.equities import (
        equity_sell_market,
        equity_buy_to_cover_market,
    )
    from schwab.orders.common import (
        OrderType,
        EquityInstruction,
        Duration,
        Session,
        OrderStrategyType,
    )
    from schwab.orders.generic import OrderBuilder
except Exception as e:  # pragma: no cover - import-time failure
    equity_sell_market = None  # type: ignore[assignment]
    equity_buy_to_cover_market = None  # type: ignore[assignment]
    OrderType = None  # type: ignore[assignment]
    EquityInstruction = None  # type: ignore[assignment]
    Duration = None  # type: ignore[assignment]
    Session = None  # type: ignore[assignment]
    OrderStrategyType = None  # type: ignore[assignment]
    OrderBuilder = None  # type: ignore[assignment]
    SCHWAB_IMPORT_ERROR = e

# ---------------------------------------------------------------------------
# Paths / config
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
LIVE_INFO_SHEET = "Daily_Trades"

DRY_RUN = False  # When True, only print planned exits; do not send orders.


def normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    s = str(direction_cell).strip().lower()
    if s in ("long", "short"):
        return s
    return s


def _tos_exit_cell_to_order_type(cell_value) -> str:
    """
    Column E: ToS Exit.
    'Open' (case-insensitive) -> 'MKT', else -> 'MOC'.
    """
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def _last_row_by_ticker(sheet, col_letter: str = "A", min_row: int = 2, max_scan: int = 5000) -> int:
    """
    Robustly compute last row by scanning a column range once and finding the last non-empty ticker.

    This avoids relying on `sheet.used_range.last_cell.row`, which can be wrong in some cases
    and causes us to miss staged trades.
    """
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


def read_exit_trade_info(sheet) -> List[Tuple[str, str, int, str]]:
    """
    Read columns A–C and E from sheet (Prices).

    Returns exits = [(ticker, action, size, order_type)], where:
      - action: 'SELL' for long, 'BUY' for short (we map BUY to buy-to-cover)
      - order_type: 'MKT' or 'MOC' based on column E (ToS Exit).
    """
    max_row = _last_row_by_ticker(sheet, col_letter="A", min_row=2, max_scan=5000)

    exits: List[Tuple[str, str, int, str]] = []

    for row in range(2, max_row + 1):
        ticker_cell = sheet.range(f"A{row}").value
        direction_cell = sheet.range(f"B{row}").value
        size_cell = sheet.range(f"C{row}").value
        tos_exit_cell = sheet.range(f"E{row}").value

        if ticker_cell is None or str(ticker_cell).strip() == "":
            continue

        ticker = str(ticker_cell).strip().upper()
        direction_norm = normalize_direction(direction_cell)

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

        order_type = _tos_exit_cell_to_order_type(tos_exit_cell)
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
        print("\nDRY_RUN is True: no Schwab exit orders will be sent. "
              "Set DRY_RUN = False at the top of this script to send live orders.")
        return

    print("\nPlacing Schwab exit orders...")
    for ticker, action, size, order_type in exits:
        try:
            # Map action string to EquityInstruction
            if action == "SELL":
                instr = EquityInstruction.SELL
            else:  # BUY to close short
                instr = EquityInstruction.BUY_TO_COVER

            # Map our string order_type to Schwab OrderType
            if order_type == "MKT":
                ot = OrderType.MARKET
            else:
                ot = OrderType.MARKET_ON_CLOSE

            ob = (
                OrderBuilder()
                .set_order_type(ot)
                .set_duration(Duration.DAY)
                .set_session(Session.NORMAL)
                .set_order_strategy_type(OrderStrategyType.SINGLE)
                .add_equity_leg(instr, ticker, size)
            )
            order_spec = ob.build()
            resp = client.place_order(account_id, order_spec)
            status = getattr(resp, "status_code", None)
            text = getattr(resp, "text", None)
            print(f"Submitted {action} {size} {ticker} ({order_type}), response: {status}")
            print(status, text)
        except Exception as e:
            print(f"Error placing Schwab exit order for {ticker}: {e}")


def main() -> int:
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return 1
    if not os.path.exists(LIVE_INFO_FILE):
        print(f"Live trade info file not found: {LIVE_INFO_FILE}")
        return 1

    app = None
    wb = None
    try:
        app = xw.App(visible=False)
        try:
            wb = app.books.open(os.path.abspath(LIVE_INFO_FILE))
        except Exception:
            print("Please close Live_Trade_Info")
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass
                app = None
            return 1
        # Decide which sheet to read from by reading Trade_Mode!C3
        # directly from the already-open xlwings workbook (avoids helper re-opening/locking issues).
        try:
            trade_mode_value = wb.sheets[TRADE_MODE_SHEET].range(TRADE_MODE_CELL).value
        except Exception:
            trade_mode_value = None

        mode_lower = str(trade_mode_value).strip().lower() if trade_mode_value is not None else ""
        full_auto = mode_lower == MODE_FULL_AUTO.lower()
        if mode_lower == MODE_SINGLE_DAY.lower():
            sheet_name = "Daily_Trades"
        elif full_auto:
            sheet_name = get_trade_sheet_name(LIVE_INFO_FILE)
            if not sheet_name:
                print("Full_Auto: today is not a trading day; nothing to exit.")
                wb.close()
                return 0
        else:
            # Emergency Exit Order Run
            # To run a specific day from the past, to send exit orders, change the line below to that day's name
            sheet_name = datetime.now().strftime("%A")
        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            print(f"Sheet '{sheet_name}' not found in {LIVE_INFO_FILE}.")
            wb.close()
            # Full_Auto with no trades for today is expected — clean success.
            return 0 if full_auto else 1

        # Refresh exit-type values from Latest Earnings (just before sending orders).
        # Latest Earnings -> Live_Trade_Info:
        #   AA (ToS Exit) -> Live_Trade_Info column E
        single_day_mode = sheet_name == "Daily_Trades"
        if full_auto:
            target_sheet_names = [sheet_name]
        else:
            target_sheet_names = ["Daily_Trades"] if single_day_mode else ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        try:
            print("\nRefreshing ToS exit types from Latest Earnings Document (AA)...")
            lookup = read_exit_types_from_latest_earnings(
                single_day=single_day_mode, full_auto=full_auto
            )
        except Exception as e:
            print(f"Failed to refresh exit types from Latest Earnings: {e}")
            wb.close()
            return 1

        updated_rows = 0
        for target_sheet_name in target_sheet_names:
            try:
                ws_target = wb.sheets[target_sheet_name]
            except Exception:
                continue
            max_row = _last_row_by_ticker(ws_target, col_letter="A", min_row=2, max_scan=5000)
            for row in range(2, max_row + 1):
                ticker_cell = ws_target.range(f"A{row}").value
                if ticker_cell is None or str(ticker_cell).strip() == "":
                    continue
                ticker = str(ticker_cell).strip().upper()
                if ticker not in lookup:
                    continue
                tos_exit_val = lookup[ticker]["tos_exit"]
                ws_target.range(f"E{row}").value = tos_exit_val
                updated_rows += 1

        if updated_rows:
            print(f"Updated {updated_rows} row(s) in Live_Trade_Info column E from Latest Earnings.")

        exits = read_exit_trade_info(sheet)

        if not exits:
            if updated_rows and not single_day_mode:
                print(
                    f"No valid exit rows on today's sheet ({sheet_name}); nothing to exit. "
                    f"(Column E was refreshed on {updated_rows} row(s) across weekday sheets.)"
                )
            elif updated_rows:
                print(
                    "No valid exit rows on Daily_Trades; nothing to exit. "
                    f"(Column E was refreshed on {updated_rows} row(s).)"
                )
            else:
                print("No valid rows in Live_Trade_Info; nothing to exit.")
            if updated_rows:
                wb.save()
            wb.close()
            return 0

        # No interactive prompt for Task Scheduler runs.
        if DRY_RUN:
            print("DRY_RUN is True: not sending Schwab exit orders.")
            if updated_rows:
                wb.save()
            wb.close()
            return 0

        try:
            client, cfg = create_client()
        except Exception as e:
            print(f"Failed to create Schwab client: {e}")
            wb.close()
            return 1

        account_id = cfg.get("account_id")
        if not account_id:
            print("account_id is missing from schwab_config.json; cannot place Schwab exit orders.")
            wb.close()
            return 1

        place_exit_orders_schwab(client, account_id, exits)

        wb.save()
        wb.close()
        return 0
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())

