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
