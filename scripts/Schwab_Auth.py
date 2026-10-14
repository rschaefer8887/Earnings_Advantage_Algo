"""
Schwab_Auth — OAuth helper and client factory for Schwab Trader API (sandbox or live).

Usage:
- Copy schwab_config.example.json to schwab_config.json at project root.
- Fill in:y
    - api_key: your Schwab app key
    - app_secret: your Schwab app secret
    - callback_url: your registered callback URL (e.g. https://127.0.0.1:8182)
    - token_path: where to store OAuth tokens (e.g. schwab_tokens.json)
    - account_id: Schwab account number/hash to trade against
- Run this script once to complete OAuth and cache tokens:
    python scripts/Schwab_Auth.py

Other scripts (e.g. Open_Trades_ToS.py) can import create_client()
to reuse the same credentials and token store.
"""

import json
import os
import asyncio
from datetime import datetime, timezone
from typing import Any, Dict, Optional, Tuple
from zoneinfo import ZoneInfo

MOUNTAIN_TZ = ZoneInfo("America/Denver")


def _format_mountain_12h(dt: datetime) -> str:
    """Format datetime in US Mountain Time, 12-hour with AM/PM (no leading zero on hour)."""
    mt = dt.astimezone(MOUNTAIN_TZ)
    hour = int(mt.strftime("%I"))  # 01-12, strip leading zero
    return mt.strftime(f"%b %d, %Y {hour}:%M %p MT")

# Refresh token expires 7 days from receipt; we warn after 6 days.
REFRESH_TOKEN_DAYS_VALID = 7
WARN_AFTER_DAYS = 6

try:
    asyncio.get_event_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

CONFIG_FILENAME = "schwab_config.json"
AUTH_TIMESTAMP_FILENAME = "schwab_auth_timestamp.json"

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
CONFIG_PATH = os.path.join(_BASE_DIR, CONFIG_FILENAME)
AUTH_TIMESTAMP_PATH = os.path.join(_BASE_DIR, AUTH_TIMESTAMP_FILENAME)

SCHWAB_PY_IMPORT_ERROR = None
try:
    from schwab.auth import client_from_token_file, easy_client
except Exception as e:  # pragma: no cover - import-time failure
    client_from_token_file = None  # type: ignore[assignment]
    easy_client = None  # type: ignore[assignment]
    SCHWAB_PY_IMPORT_ERROR = e


def read_auth_timestamp(path: str = AUTH_TIMESTAMP_PATH) -> Tuple[Optional[datetime], Optional[float], Optional[str]]:
    """
    Read the last-auth timestamp from disk.
    Returns (timestamp_utc, age_in_days, display_mountain) or (None, None, None) if missing/invalid.
    display_mountain is 12-hour Mountain Time string for printing.
    """
    if not os.path.exists(path):
        return None, None, None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        ts_str = data.get("last_auth_utc")
        if not ts_str:
            return None, None, None
        ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        age = (datetime.now(timezone.utc) - ts).total_seconds() / (24 * 3600)
        display = data.get("last_auth_mountain") or _format_mountain_12h(ts)
        return ts, age, display
    except (json.JSONDecodeError, ValueError, OSError):
        return None, None, None


def write_auth_timestamp(path: str = AUTH_TIMESTAMP_PATH) -> None:
    """Record current time when a new refresh token was created (full OAuth). Stores UTC and Mountain Time (12-hour)."""
    now_utc = datetime.now(timezone.utc)
    data = {
        "last_auth_utc": now_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "last_auth_mountain": _format_mountain_12h(now_utc),
    }
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=2)
        f.write("\n")


def print_account_numbers_for_config(client: Any, cfg: Dict[str, Any]) -> None:
    """
    Print accountNumber / hashValue for each linked account (API hashes for place_order / account_id).
    Output is sensitive; use locally only.
    """
    try:
        resp = client.get_account_numbers()
    except Exception as e:
        print(f"Could not fetch account numbers: {e}")
        return
    if getattr(resp, "status_code", None) != 200:
        text = getattr(resp, "text", "") or ""
        print(f"get_account_numbers failed: {resp.status_code} {text[:500]}")
        return
    try:
        rows = resp.json()
    except Exception as e:
        print(f"Could not parse get_account_numbers response: {e}")
        return
    print("\n--- Linked accounts: accountNumber → hashValue ---")
    print("(Put hashValue in schwab_config.json as account_id for order placement.)\n")
    print(json.dumps(rows, indent=2))
    configured = cfg.get("account_id")
    if configured and isinstance(rows, list):
        for row in rows:
            if not isinstance(row, dict):
                continue
            if row.get("hashValue") == configured:
                acct = row.get("accountNumber", "?")
                print(f"\nConfigured account_id matches accountNumber {acct}.")
                break


def check_expiry_warning(path: str = AUTH_TIMESTAMP_PATH) -> None:
    """
    If a timestamp exists and is >= WARN_AFTER_DAYS old, print a warning
    that the refresh token will expire soon (7-day validity).
    """
    _, age, display_mt = read_auth_timestamp(path)
    if age is None:
        return
    if age >= WARN_AFTER_DAYS:
        days_left = max(0, REFRESH_TOKEN_DAYS_VALID - age)
        last_str = f" (Last auth: {display_mt})" if display_mt else ""
        print(
            f"\n*** WARNING: Schwab refresh token is approaching expiry. "
            f"Last auth was {age:.1f} days ago{last_str}; token is valid for {REFRESH_TOKEN_DAYS_VALID} days. "
            f"Approx. {days_left:.1f} day(s) remaining. Re-run this script to re-authenticate. ***\n"
        )


def load_config(path: str = CONFIG_PATH) -> Dict[str, Any]:
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Schwab config file not found: {path}\n"
            "Copy schwab_config.example.json to schwab_config.json"
            " and fill in your api_key, app_secret, callback_url, token_path, account_id."
        )
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    required = ["api_key", "app_secret", "callback_url", "token_path", "account_id"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(f"Missing Schwab config keys: {', '.join(missing)} in {path}")
    return cfg


def create_client() -> Tuple[Any, Dict[str, Any]]:
    """
    Create and return a Schwab API client plus the loaded config.

    Returns:
        (client, cfg) where client is the easy_client result, cfg is the config dict.
    """
    if easy_client is None:
        raise ImportError(
            "Could not import schwab-py. Install it with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_PY_IMPORT_ERROR}"
        )
    cfg = load_config()
    client = easy_client(
        api_key=cfg["api_key"],
        app_secret=cfg["app_secret"],
        callback_url=cfg["callback_url"],
        token_path=os.path.join(_BASE_DIR, cfg["token_path"]),
        requested_browser="windows-default",
        interactive=False,
    )
    return client, cfg


def create_client_from_token() -> Tuple[Any, Dict[str, Any]]:
    """
    Create a Schwab API client from the cached token file only.

    Does not start browser OAuth or delete expired tokens. Run Schwab_Auth.py
    first if authentication is needed.
    """
    if client_from_token_file is None:
        raise ImportError(
            "Could not import schwab-py. Install it with:\n"
            "    python -m pip install --upgrade schwab-py\n"
            f"Underlying import error: {SCHWAB_PY_IMPORT_ERROR}"
        )
    cfg = load_config()
    token_path = os.path.join(_BASE_DIR, cfg["token_path"])
    if not os.path.isfile(token_path):
        raise FileNotFoundError(
            f"Schwab token file not found: {token_path}\n"
            "Run python scripts/Schwab_Auth.py first to authenticate."
        )
    client = client_from_token_file(
        token_path,
        cfg["api_key"],
        cfg["app_secret"],
    )
    return client, cfg


def main() -> None:
    """
    Create a client and print basic account info.
    Asks whether to do a full OAuth reauthorization (resets the 7-day refresh-token clock).
    Otherwise uses existing refresh token if present.
    On success, prints get_account_numbers() (accountNumber and hashValue) for schwab_config.
    """
    cfg = load_config()
    token_path_full = os.path.join(_BASE_DIR, cfg["token_path"])
    token_file_existed = os.path.exists(token_path_full)

    did_full_oauth = False
    if token_file_existed:
        reply = input("Do a full OAuth reauthorization? (resets 7-day clock) [y/N]: ").strip().lower()
        if reply in ("y", "yes"):
            os.remove(token_path_full)
            did_full_oauth = True
            print("Full OAuth reauthorization: token file removed; browser will open to sign in.")
    else:
        did_full_oauth = True  # first run: no token file, OAuth will run

    try:
        client, cfg = create_client()
    except Exception as e:
        print(e)
        return

    check_expiry_warning()

    try:
        resp = client.get_accounts()
        try:
            data = resp.json()
        except Exception:
            data = None
        if isinstance(data, list):
            count = len(data)
        elif data:
            count = 1
        else:
            count = 0
    except Exception as e:
        print(f"Authenticated but failed to fetch accounts: {e}")
        return

    if did_full_oauth:
        write_auth_timestamp()
    print("Schwab authentication successful.")
    print(f"Loaded {count} account record(s). Configured account_id: {cfg.get('account_id')}")
    print_account_numbers_for_config(client, cfg)


if __name__ == "__main__":
    main()
