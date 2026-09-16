"""
Shared helper for Live_Trade_Info.xlsx: resolve which sheet to use for trade data.

Reads Trade_Mode!C3:
  - "Single Day" -> Daily_Trades
  - "Full_Auto"  -> date sheet for today (e.g. 9_10_Th); falls back to Daily_Trades
                   if today is not a trading day
  - otherwise    -> current weekday name (Monday, ...)

If the file or sheet is missing, or the cell cannot be read, returns "Daily_Trades".
"""

import os
from datetime import datetime
from typing import Any, Optional

TRADE_MODE_SHEET = "Trade_Mode"
TRADE_MODE_CELL = "C3"
MODE_SINGLE_DAY = "Single Day"
MODE_FULL_AUTO = "Full_Auto"
DEFAULT_SHEET = "Daily_Trades"


def _read_trade_mode_cell(file_path: str) -> Any:
    openpyxl_value = None
    try:
        from openpyxl import load_workbook

        wb = load_workbook(file_path, data_only=True)
        try:
            if TRADE_MODE_SHEET not in wb.sheetnames:
                openpyxl_value = None
            else:
                openpyxl_value = wb[TRADE_MODE_SHEET][TRADE_MODE_CELL].value
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        openpyxl_value = None

    value = openpyxl_value
    if value is None:
        try:
            import xlwings as xw  # type: ignore

            app = xw.App(visible=False)
            wb = None
            try:
                wb = app.books.open(file_path)
                value = wb.sheets[TRADE_MODE_SHEET].range(TRADE_MODE_CELL).value
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass
                try:
                    app.quit()
                except Exception:
                    pass
        except Exception:
            value = None
    return value


def get_trade_sheet_name(file_path: str) -> str:
    """
    Return the Live_Trade_Info sheet name to use for reading or writing trade data.
    """
    if not file_path or not os.path.exists(file_path):
        return DEFAULT_SHEET

    value = _read_trade_mode_cell(file_path)
    if value is None:
        return DEFAULT_SHEET

    low = str(value).strip().lower()
    if low == MODE_SINGLE_DAY.lower():
        return DEFAULT_SHEET

    if low == MODE_FULL_AUTO.lower():
        from full_auto_utils import full_auto_sheet_name_for_today

        name = full_auto_sheet_name_for_today()
        if name is None:
            print(
                "Full_Auto: today is not a trading day; no date sheet to use "
                f"(defaulting unavailable — callers should handle missing sheet)."
            )
            return ""
        return name

    return datetime.now().strftime("%A")


def trade_mode_is_full_auto(file_path: str) -> bool:
    if not file_path or not os.path.exists(file_path):
        return False
    value = _read_trade_mode_cell(file_path)
    return str(value or "").strip().lower() == MODE_FULL_AUTO.lower()


def live_info_has_trades_for_day(file_path: str, day) -> bool:
    """
    True if Live_Trade_Info has at least one valid trade row on the date sheet for day
    (Full_Auto), or on the sheet get_trade_sheet_name would use when day is today.
    """
    return bool(read_live_trades_for_day(file_path, day))


def read_live_trades_for_day(file_path: str, day) -> list:
    """
    Read Live_Trade_Info date sheet for day.

    Returns list of (ticker, direction) for valid rows (size may still be formula/empty).
    """
    if not file_path or not os.path.exists(file_path):
        return []
    try:
        from full_auto_utils import dow_abbrev_for_date, sheet_name_for_trade_date
        from openpyxl import load_workbook

        sheet_name = sheet_name_for_trade_date(day, dow_abbrev_for_date(day))
        wb = load_workbook(file_path, data_only=True)
        try:
            if sheet_name not in wb.sheetnames:
                return []
            ws = wb[sheet_name]
            out = []
            for row in range(2, (ws.max_row or 1) + 1):
                ticker_raw = ws.cell(row=row, column=1).value
                if ticker_raw is None or str(ticker_raw).strip() == "":
                    continue
                ticker = str(ticker_raw).strip().upper()
                direction = str(ws.cell(row=row, column=2).value or "").strip().lower()
                if direction not in ("long", "short"):
                    continue
                out.append((ticker, direction))
            return out
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        return []


def write_live_trades_for_day(
    file_path: str,
    day,
    trades: list,
) -> str:
    """
    Write/replace the date sheet for day with trades.

    trades: sequence of (ticker, direction, size, ibkr_exit="", tos_exit="")
    Returns the sheet name written.
    """
    from full_auto_utils import dow_abbrev_for_date, sheet_name_for_trade_date
    from openpyxl import Workbook, load_workbook

    sheet_name = sheet_name_for_trade_date(day, dow_abbrev_for_date(day))
    if not os.path.exists(file_path):
        wb = Workbook()
        wb.remove(wb.active)
    else:
        wb = load_workbook(file_path)

    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]

    # Clear existing used area (keep simple: overwrite from row 1)
    if ws.max_row and ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    ws["A1"] = "Ticker"
    ws["B1"] = "Direction"
    ws["C1"] = "Share Size"
    ws["D1"] = "IBKR Exit"
    ws["E1"] = "ToS Exit"

    for i, item in enumerate(trades, start=2):
        ticker, direction, size = item[0], item[1], item[2]
        ibkr = item[3] if len(item) > 3 else ""
        tos = item[4] if len(item) > 4 else ""
        ws.cell(row=i, column=1).value = ticker
        ws.cell(row=i, column=2).value = direction
        ws.cell(row=i, column=3).value = size
        ws.cell(row=i, column=4).value = ibkr or None
        ws.cell(row=i, column=5).value = tos or None

    wb.save(file_path)
    try:
        wb.close()
    except Exception:
        pass
    return sheet_name
