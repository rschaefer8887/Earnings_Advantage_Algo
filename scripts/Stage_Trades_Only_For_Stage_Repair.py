"""
Stage_Trades_Only_For_Stage_Repair — Full_Auto stage without price flagging.

Same rules as Stage_Trades_Auto in Full_Auto:
  - Stage from first trade-row C (else first M1) downward
  - Columns M + K + Y + AC (+ Z size, AA/AB exits) → Live_Trade_Info date sheets

Does NOT run flag_prices_full_auto and does NOT save Latest Earnings P/Q flags.

Usage:
  python scripts/Stage_Trades_Only_For_Stage_Repair.py
"""

from __future__ import annotations

import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from Stage_Trades_Auto import (  # noqa: E402
    SOURCE_FILE,
    SOURCE_SHEET,
    _collect_full_auto_trades_xw,
    _write_live_trade_info_sheets,
)
from earnings_workbook_utils import (  # noqa: E402
    open_or_attach_earnings_workbook,
    release_earnings_workbook,
)
from full_auto_utils import MODE_FULL_AUTO, is_full_auto_mode, read_trade_mode_value  # noqa: E402
from run_logging import install_run_log  # noqa: E402

_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
OUTPUT_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")


def main() -> int:
    install_run_log("Stage_Trades_Only_For_Stage_Repair")

    if not os.path.exists(SOURCE_FILE):
        print(f"ERROR: Source Earnings file not found: {SOURCE_FILE}")
        return 1

    mode_raw = read_trade_mode_value(OUTPUT_FILE) if os.path.exists(OUTPUT_FILE) else None
    if not is_full_auto_mode(mode_raw):
        print(
            f"ERROR: Trade_Mode is {mode_raw!r}; this repair script only supports "
            f"{MODE_FULL_AUTO}."
        )
        return 1

    print(
        f"Trade_Mode is {MODE_FULL_AUTO}: stage-only repair "
        "(no flag_prices_full_auto; Earnings P/Q untouched)."
    )

    app = None
    wb = None
    owned_app = False
    owned_book = False
    try:
        try:
            app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(
                SOURCE_FILE
            )
        except Exception as e:
            print(f"FAIL: could not open/attach Latest Earnings: {e}")
            return 1

        if owned_book:
            print("Opened Latest Earnings in hidden Excel (read-only for staging).")
        else:
            print("Attached to Latest Earnings already open in Excel.")

        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"ERROR: Worksheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
            return 1

        by_sheet = _collect_full_auto_trades_xw(sheet)
        if not by_sheet:
            print(
                "FAIL: no eligible trades to stage (from C/M1 start downward)."
            )
            return 1

        try:
            _write_live_trade_info_sheets(by_sheet)
        except PermissionError:
            print(
                "FAIL: could not update Live_Trade_Info "
                "(file may be open). Close it and retry."
            )
            return 1

        print("Stage-only repair complete (flagging skipped).")
        return 0
    except Exception as e:
        print(f"FAIL: Stage_Trades_Only_For_Stage_Repair crashed: {e}")
        return 1
    finally:
        release_earnings_workbook(
            app,
            wb,
            owned_app=owned_app,
            owned_book=owned_book,
            earnings_file=SOURCE_FILE,
        )


if __name__ == "__main__":
    raise SystemExit(main())
