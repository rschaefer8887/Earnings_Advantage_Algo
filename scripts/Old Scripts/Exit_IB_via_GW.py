"""
Exit Live Trades via IB Gateway with optional Open → MOC changes in Live_Trade_Info.xlsx

Same as Exit_GW plus: before reading exit types, asks if any symbols should have their
order type changed from "Open" to "MOC". If yes, lists symbols, you enter which ones
(comma-space separated, up to 10); script writes "MOC" to column D for those rows,
saves the workbook, then reads column D and sends MOC/MKT orders accordingly.

- Reads/writes Live_Trade_Info.xlsx (sheet "Daily_Trades"), columns A–D via xlwings.
- Column D: "Open" → MKT, else → MOC. You can change Open → MOC before sending.
- Sends IB exit orders without interactive prompts (Task Scheduler friendly).

Prerequisites: IB Gateway running (API enabled), pip install ib_insync xlwings.
Close Live_Trade_Info.xlsx in Excel before running.
"""

import os
import asyncio
import warnings
from typing import List, Tuple

warnings.filterwarnings("ignore", message=".*Unknown extension.*", category=UserWarning)
warnings.filterwarnings(
    "ignore", message=".*Conditional Formatting extension.*", category=UserWarning
)

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

IB_IMPORT_ERROR = None
try:
    from ib_insync import IB, Stock, Order
except Exception as e:
    IB = None  # type: ignore[assignment]
    IB_IMPORT_ERROR = e

try:
    import xlwings as xw
except ImportError:
    xw = None

from datetime import datetime
import sys

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # scripts/Old Scripts
_SCRIPTS_DIR = os.path.dirname(_SCRIPT_DIR)  # scripts/
_BASE_DIR = os.path.dirname(_SCRIPTS_DIR)  # repo root

# live_trade_info_utils lives in scripts/; this util lives beside this file.
for _p in (_SCRIPT_DIR, _SCRIPTS_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from live_trade_info_utils import MODE_SINGLE_DAY, TRADE_MODE_CELL, TRADE_MODE_SHEET
from earnings_exit_type_utils import read_exit_types_from_latest_earnings

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
LIVE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")
LIVE_INFO_SHEET = "Daily_Trades"

IB_HOST = "127.0.0.1"
IB_PORT = 4001
IB_CLIENT_ID = 4
IB_ACCOUNT = "U24159961"
DEFAULT_EXCHANGE = "SMART"
DEFAULT_CURRENCY = "USD"
DRY_RUN = False

MAX_SYMBOLS_TO_CHANGE = 10


def normalize_direction(direction_cell) -> str:
    if direction_cell is None:
        return ""
    s = str(direction_cell).strip().lower()
    if s in ("long", "short"):
        return s
    return s


def _exit_type_cell_to_order_type(cell_value) -> str:
    if cell_value is None or not str(cell_value).strip():
        return "MOC"
    if str(cell_value).strip().lower() == "open":
        return "MKT"
    return "MOC"


def _get_symbols_from_sheet(sheet) -> List[str]:
    """Return list of non-empty tickers from column A (row 2 onward), order preserved."""
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000
    symbols: List[str] = []
    for row in range(2, max_row + 1):
        cell = sheet.range(f"A{row}").value
        if cell is None or str(cell).strip() == "":
            continue
        symbols.append(str(cell).strip().upper())
    return symbols


def _set_exit_type_to_moc(sheet, tickers: List[str]) -> int:
    """
    For each ticker in tickers, find row(s) in column A where value matches (case-insensitive)
    and set column D to "MOC". Returns number of cells updated.
    """
    try:
        max_row = sheet.used_range.last_cell.row
    except Exception:
        max_row = 1000
    ticker_set = {t.upper() for t in tickers}
    updated = 0
    for row in range(2, max_row + 1):
        cell = sheet.range(f"A{row}").value
        if cell is None or str(cell).strip() == "":
            continue
        if str(cell).strip().upper() in ticker_set:
            sheet.range(f"D{row}").value = "MOC"
            updated += 1
    return updated


def _parse_moc_input(user_input: str) -> List[str]:
    """Parse 'AAPL, WMT, MSFT' into ['AAPL','WMT','MSFT'], max 10, strip and uppercase."""
    parts = [p.strip().upper() for p in user_input.split(",") if p.strip()]
    return parts[:MAX_SYMBOLS_TO_CHANGE]


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


def read_exit_trade_info(sheet):
    """
    Read columns A–D from sheet. Column D = IB Exit.
    Returns exits = [(ticker, action, size, order_type), ...].
    """
    max_row = _last_row_by_ticker(sheet, col_letter="A", min_row=2, max_scan=5000)

    exits: List[Tuple[str, str, int, str]] = []

    for row in range(2, max_row + 1):
        ticker_cell = sheet.range(f"A{row}").value
        direction_cell = sheet.range(f"B{row}").value
        size_cell = sheet.range(f"C{row}").value
        exit_type_cell = sheet.range(f"D{row}").value

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

        order_type = _exit_type_cell_to_order_type(exit_type_cell)
        action = "SELL" if direction_norm == "long" else "BUY"
        exits.append((ticker, action, size, order_type))

    return exits


def connect_ib() -> IB:
    if IB is None:
        raise ImportError(
            "Could not import ib_insync or its dependencies. "
            f"Details:\n    {IB_IMPORT_ERROR}\n\n"
            "Try: python -m pip install --upgrade ib_insync eventkit nest-asyncio numpy"
        )
    ib = IB()
    print(f"Connecting to IB Gateway at {IB_HOST}:{IB_PORT} with clientId={IB_CLIENT_ID} ...")
    ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
    print("Connected to IB Gateway.")
    return ib


def place_exit_orders_ib(ib: IB, exits: List[Tuple[str, str, int, str]]) -> None:
    if not exits:
        print("No exit orders to place.")
        return

    print("\nPlanned exit orders (close/cover):")
    for ticker, action, size, order_type in exits:
        print(f"  {action} {size} {ticker}  [{order_type}]")

    if DRY_RUN:
        print("\nDRY_RUN is True: no orders will be sent. Set DRY_RUN = False to send exit orders.")
        return

    print("\nPlacing exit orders...")
    for ticker, action, size, order_type in exits:
        contract = Stock(ticker, DEFAULT_EXCHANGE, DEFAULT_CURRENCY)
        order = Order(
            action=action,
            orderType=order_type,
            totalQuantity=size,
            tif="DAY",
        )
        if IB_ACCOUNT:
            order.account = IB_ACCOUNT
        trade = ib.placeOrder(contract, order)
        print(f"Submitted {action} {size} {ticker} ({order_type}), orderId={trade.order.orderId}")

    ib.sleep(2)
    print("\nOrder statuses:")
    for t in ib.trades():
        print(
            f"  orderId={t.order.orderId} status={t.orderStatus.status} "
            f"filled={t.orderStatus.filled} remaining={t.orderStatus.remaining}"
        )


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

        if trade_mode_value is not None and str(trade_mode_value).strip().lower() == MODE_SINGLE_DAY.lower():
            sheet_name = "Daily_Trades"
        else:
            sheet_name = datetime.now().strftime("%A")
        try:
            sheet = wb.sheets[sheet_name]
        except Exception:
            print(f"Sheet '{sheet_name}' not found in {LIVE_INFO_FILE}.")
            wb.close()
            return 1

        # Refresh exit-type values from Latest Earnings (just before sending orders).
        # Latest Earnings -> Live_Trade_Info:
        #   AB (IB Exit) -> Live_Trade_Info column D
        single_day_mode = sheet_name == "Daily_Trades"
        # Only update the single sheet we are about to read/place orders from.
        # (Multi-day mode still filters earnings by 1..5, but we only write back to the current sheet.)
        target_sheet_names = [sheet_name]
        try:
            print("\nRefreshing IB exit types from Latest Earnings Season (AB)...")
            lookup = read_exit_types_from_latest_earnings(single_day=single_day_mode)
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
            if max_row < 2:
                continue

            tickers_col = ws_target.range(f"A2:A{max_row}").value
            existing_d_col = ws_target.range(f"D2:D{max_row}").value

            # xlwings returns a 2D list for multi-cell ranges: [[val],[val],...]
            tickers_1d = []
            if tickers_col is None:
                continue
            if isinstance(tickers_col, list):
                for item in tickers_col:
                    if isinstance(item, list):
                        tickers_1d.append(item[0])
                    else:
                        tickers_1d.append(item)
            else:
                tickers_1d = [tickers_col]

            existing_d_1d = []
            if existing_d_col is None:
                existing_d_1d = [None] * len(tickers_1d)
            elif isinstance(existing_d_col, list):
                for item in existing_d_col:
                    if isinstance(item, list):
                        existing_d_1d.append(item[0])
                    else:
                        existing_d_1d.append(item)
            else:
                existing_d_1d = [existing_d_col]

            new_d_1d = []
            for idx, ticker_cell in enumerate(tickers_1d):
                if ticker_cell is None or str(ticker_cell).strip() == "":
                    new_d_1d.append(existing_d_1d[idx] if idx < len(existing_d_1d) else None)
                    continue
                ticker = str(ticker_cell).strip().upper()
                if ticker not in lookup:
                    new_d_1d.append(existing_d_1d[idx] if idx < len(existing_d_1d) else None)
                    continue
                new_d_1d.append(lookup[ticker]["ib_exit"])
                updated_rows += 1

            # Write back as a column range in one operation.
            ws_target.range(f"D2:D{max_row}").value = [[v] for v in new_d_1d]

        if updated_rows:
            print(f"Updated {updated_rows} row(s) in Live_Trade_Info column D from Latest Earnings.")

        exits = read_exit_trade_info(sheet)

        if not exits:
            if updated_rows:
                print(
                    f"No valid exit rows on sheet '{sheet_name}'; nothing to exit. "
                    f"(Column D was refreshed on {updated_rows} row(s) from Latest Earnings.)"
                )
            else:
                print("No valid rows in Live_Trade_Info; nothing to exit.")
            if updated_rows:
                wb.save()
            wb.close()
            return 0

        if DRY_RUN:
            print("DRY_RUN is True: not sending IB exit orders.")
            if updated_rows:
                wb.save()
            wb.close()
            return 0

        try:
            ib = connect_ib()
        except Exception as e:
            print(f"Failed to connect to IB Gateway: {e}")
            wb.close()
            return 1

        try:
            place_exit_orders_ib(ib, exits)
        finally:
            print("Disconnecting from IB Gateway...")
            ib.disconnect()
            print("Disconnected.")

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
