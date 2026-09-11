"""Success timestamps for Get_Opens_ToS / Get_Closes_ToS (Full_Auto evening QC)."""

from __future__ import annotations

import json
import os
from datetime import date, datetime, timezone
from typing import Optional, Tuple
from zoneinfo import ZoneInfo

MOUNTAIN_TZ = ZoneInfo("America/Denver")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

OPENS_TIMESTAMP_PATH = os.path.join(_BASE_DIR, "get_opens_last_success.json")
CLOSES_TIMESTAMP_PATH = os.path.join(_BASE_DIR, "get_closes_last_success.json")


def _format_mountain_12h(dt: datetime) -> str:
    mt = dt.astimezone(MOUNTAIN_TZ)
    hour = int(mt.strftime("%I"))
    return mt.strftime(f"%b %d, %Y {hour}:%M %p MT")


def write_price_job_success(path: str, job_name: str) -> None:
    now_utc = datetime.now(timezone.utc)
    data = {
        "job": job_name,
        "last_success_utc": now_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "last_success_mountain": _format_mountain_12h(now_utc),
        "last_success_date_mt": now_utc.astimezone(MOUNTAIN_TZ).date().isoformat(),
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    print(f"Recorded {job_name} success: {data['last_success_mountain']}")


def write_opens_success() -> None:
    write_price_job_success(OPENS_TIMESTAMP_PATH, "Get_Opens_ToS")


def write_closes_success() -> None:
    write_price_job_success(CLOSES_TIMESTAMP_PATH, "Get_Closes_ToS")


def read_price_job_success_date_mt(path: str) -> Optional[date]:
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        raw = data.get("last_success_date_mt")
        if raw:
            return date.fromisoformat(str(raw))
        ts_str = data.get("last_success_utc")
        if not ts_str:
            return None
        ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        return ts.astimezone(MOUNTAIN_TZ).date()
    except (json.JSONDecodeError, ValueError, OSError, TypeError):
        return None


def opens_and_closes_succeeded_on(day: date) -> Tuple[bool, Optional[date], Optional[date]]:
    """Return (both_ok, opens_date_mt, closes_date_mt)."""
    o = read_price_job_success_date_mt(OPENS_TIMESTAMP_PATH)
    c = read_price_job_success_date_mt(CLOSES_TIMESTAMP_PATH)
    return (o == day and c == day), o, c