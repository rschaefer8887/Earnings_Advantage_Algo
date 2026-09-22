"""
Exit_ToS2 — Exit live trades via Schwab (secondary account) from Live_Trade_Info.xlsx.

Uses column D (IBKR Exit staging): "Open" → MARKET, else → MARKET_ON_CLOSE.
Uses account_id2 from schwab_config.json.
Requires Live_Trade_Info closed in Excel before running.
"""

import exit_schwab_common as exit_common
from exit_schwab_common import run_exit_main

DRY_RUN = False  # True = print planned exits only; do not send orders.


if __name__ == "__main__":
    exit_common.DRY_RUN = DRY_RUN
    raise SystemExit(run_exit_main(account_key="account_id2", exit_col="D"))
