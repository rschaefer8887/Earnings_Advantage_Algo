"""
Cancel_All_ToS - Cancel open Schwab (ThinkOrSwim) orders on selected accounts.

Manual emergency backup script. Cancels pending/working orders only; does not close
filled positions.

Accounts (from schwab_config.json):
  - account_id  - primary ToS account (Open_Trades_ToS.py / Exit_ToS.py) — typical MOC
  - account_id2 - secondary ToS account (Open_Trade_ToS2.py / Exit_ToS2.py) — typical Open

Prerequisites:
  - schwab-py installed: python -m pip install --upgrade schwab-py
  - Run Schwab_Auth.py first so schwab_config.json and token file exist

Usage:
  python scripts/Cancel_All_ToS.py
  Then choose 1 (ToS), 2 (ToS2), or 3 (both) at the prompt.
"""

from typing import Any, Dict, List, Optional, Set, Tuple

from Schwab_Auth import create_client_from_token

try:
    from schwab.client import Client
except Exception as e:  # pragma: no cover - import-time failure
    Client = None  # type: ignore[assignment, misc]
    SCHWAB_IMPORT_ERROR = e
else:
    SCHWAB_IMPORT_ERROR = None

# When True, list open orders but do not send cancel requests.
DRY_RUN = False

# Schwab order statuses that are still open and may be cancelled.
CANCELLABLE_STATUSES: Set[str] = {
    Client.Order.Status.AWAITING_PARENT_ORDER.value,
    Client.Order.Status.AWAITING_CONDITION.value,
    Client.Order.Status.AWAITING_STOP_CONDITION.value,
    Client.Order.Status.AWAITING_MANUAL_REVIEW.value,
    Client.Order.Status.ACCEPTED.value,
    Client.Order.Status.AWAITING_UR_OUT.value,
    Client.Order.Status.PENDING_ACTIVATION.value,
    Client.Order.Status.QUEUED.value,
    Client.Order.Status.WORKING.value,
    Client.Order.Status.NEW.value,
    Client.Order.Status.AWAITING_RELEASE_TIME.value,
    Client.Order.Status.PENDING_ACKNOWLEDGEMENT.value,
    Client.Order.Status.PENDING_RECALL.value,
} if Client is not None else set()


def _order_summary(order: Dict[str, Any]) -> str:
    """Build a short human-readable line for an order."""
    order_id = order.get("orderId", "?")
    status = order.get("status", "?")
    order_type = order.get("orderType", "?")
    legs = order.get("orderLegCollection") or []
    leg_parts: List[str] = []
    for leg in legs:
        instr = leg.get("instrument") or {}
        symbol = instr.get("symbol", "?")
        instruction = leg.get("instruction", "?")
        qty = leg.get("quantity", "?")
        leg_parts.append(f"{instruction} {qty} {symbol}")
    legs_str = "; ".join(leg_parts) if leg_parts else "no legs"
    return f"orderId={order_id} status={status} type={order_type} {legs_str}"


def _fetch_open_orders(client: Any, account_hash: str) -> List[Dict[str, Any]]:
    """Return orders on the account that are still cancellable."""
    resp = client.get_orders_for_account(account_hash)
    status_code = getattr(resp, "status_code", None)
    if status_code != 200:
        text = getattr(resp, "text", "") or ""
        raise RuntimeError(
            f"get_orders_for_account failed ({status_code}): {text[:500]}"
        )
    orders = resp.json()
    if not isinstance(orders, list):
        raise RuntimeError(f"Unexpected orders response type: {type(orders).__name__}")

    open_orders: List[Dict[str, Any]] = []
    for order in orders:
        if not isinstance(order, dict):
            continue
        if order.get("status") in CANCELLABLE_STATUSES:
            open_orders.append(order)
    return open_orders


def _prompt_account_choice() -> str:
    """
    Ask which account(s) to cancel. Returns '1', '2', or '3'.
    """
    print("Cancel which accounts?")
    print("  1) Cancel all orders in ToS, your typical MOC account.")
    print("  2) Cancel all orders in ToS2, your typical Open account.")
    print("  3) Cancel all orders in both ToS accounts.")
    while True:
        choice = input("Enter 1-3: ").strip()
        if choice in ("1", "2", "3"):
            return choice
        print("Please enter 1, 2, or 3.")


def cancel_account_orders(
    client: Any,
    account_hash: str,
    label: str,
) -> Tuple[int, int]:
    """
    List and optionally cancel open orders for one account.

    Returns (cancelled_count, error_count).
    """
    print(f"\n--- {label} ---")
    try:
        open_orders = _fetch_open_orders(client, account_hash)
    except Exception as e:
        print(f"Failed to fetch orders: {e}")
        return 0, 1

    if not open_orders:
        print("No open orders to cancel.")
        return 0, 0

    print(f"Found {len(open_orders)} open order(s):")
    for order in open_orders:
        print(f"  {_order_summary(order)}")

    if DRY_RUN:
        print("DRY_RUN is True: no cancel requests sent.")
        return 0, 0

    cancelled = 0
    errors = 0
    for order in open_orders:
        order_id = order.get("orderId")
        if order_id is None:
            print(f"  Skipping order with no orderId: {_order_summary(order)}")
            errors += 1
            continue
        try:
            resp = client.cancel_order(order_id, account_hash)
            status_code = getattr(resp, "status_code", None)
            if status_code in (200, 201, 204):
                print(f"  Cancelled orderId={order_id} ({status_code})")
                cancelled += 1
            else:
                text = getattr(resp, "text", "") or ""
                print(f"  Failed to cancel orderId={order_id}: {status_code} {text[:300]}")
                errors += 1
        except Exception as e:
            print(f"  Error cancelling orderId={order_id}: {e}")
            errors += 1

    return cancelled, errors


def main() -> int:
    if SCHWAB_IMPORT_ERROR is not None or Client is None:
        print(
            "Could not import schwab-py. Install it with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_IMPORT_ERROR}"
        )
        return 1

    mode = "DRY RUN (list only)" if DRY_RUN else "LIVE (will cancel orders)"
    print(f"Cancel_All_ToS - {mode}")
    print("Cancels open Schwab orders only; does not close filled positions.\n")

    choice = _prompt_account_choice()

    try:
        client, cfg = create_client_from_token()
    except Exception as e:
        print(f"Failed to create Schwab client: {e}")
        return 1

    all_accounts: List[Tuple[str, Optional[str]]] = [
        ("ToS (MOC)", cfg.get("account_id")),
        ("ToS2 (Open)", cfg.get("account_id2")),
    ]
    if choice == "1":
        accounts = [all_accounts[0]]
    elif choice == "2":
        accounts = [all_accounts[1]]
    else:
        accounts = all_accounts

    if choice in ("1", "3") and not cfg.get("account_id"):
        print("account_id is missing from schwab_config.json; cannot continue.")
        return 1
    if choice in ("2", "3") and not cfg.get("account_id2"):
        print("account_id2 is missing from schwab_config.json; cannot continue.")
        return 1

    total_cancelled = 0
    total_errors = 0
    accounts_checked = 0

    for label, account_hash in accounts:
        if not account_hash:
            print(f"\n--- {label} ---")
            print("Account hash not configured; skipping.")
            continue
        accounts_checked += 1
        cancelled, errors = cancel_account_orders(client, account_hash, label)
        total_cancelled += cancelled
        total_errors += errors

    if accounts_checked == 0:
        print("No Schwab account hashes configured.")
        return 1

    print("\n--- Summary ---")
    if DRY_RUN:
        print("DRY_RUN complete. Set DRY_RUN = False to cancel for real.")
    else:
        print(f"Cancelled {total_cancelled} order(s).")
        if total_errors:
            print(f"Errors: {total_errors}")

    return 1 if total_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
