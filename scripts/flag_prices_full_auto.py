"""
Full_Auto price flagging for Latest Earnings columns P and Q.

Invoked from Stage_Trades_Auto when Trade_Mode is Full_Auto (same Task Scheduler job).

Price calendar (entry day D from M+K, trading days only):
  - today == D   -> may run C->S job for that date's groups
  - today == D+1 -> P: O ... 0 and Q: M1 at same start row (T and U)
  - today == D+2 -> Q: M2 at start (closes -> V)

Required C (always, after date jobs):
  - Morning / default: first trade row above zz with blank column T (not entered).
  - Evening + open orders today: M1+O on today's weekday block; C on the first
    trade row after that block.
  - Exception: no C when the last column-M weekday block above zz already starts
    with M1 or M2.
  - Q trailing 0 for C sits on the column-A zz end-of-trades row.
  - Evening only: next trading day's Live_Trade_Info trades — paste column Z as values
    (lock share size). Skipped on morning / same-day Stage runs.

Get_Closes_ToS Q chain: M2 until M1; M1 until C; C until 0.
  - O (column P) trailing 0 must sit on the same row as C (P=0, Q=C).

After writing flags, validate_flag_layout_strict() enforces layout invariants.
Failure raises FlagLayoutError (Stage exits non-zero).

Evening Stage (after successful Get_Opens + Get_Closes today, dests filled):
  clear completed flags and place as-of the next trading day.
Morning Stage: place as-of calendar today.

"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from typing import Any, List, Optional, Sequence

from full_auto_utils import (
    add_trading_days,
    cell_has_value,
    dow_abbrev_for_date,
    is_trading_day,
    next_trading_day,
    parse_earnings_date,
    parse_m_dow,
    resolve_trade_date,
)
from price_run_timestamps import (
    open_trades_both_succeeded_on,
    opens_and_closes_succeeded_on,
)
from live_trade_info_utils import (
    live_info_has_trades_for_day,
    read_live_trades_for_day,
    write_live_trades_for_day,
)
from earnings_workbook_utils import (
    LATEST_EARNINGS_FILE,
    LATEST_EARNINGS_SHEET,
    open_or_attach_earnings_workbook,
    release_earnings_workbook,
)

HEADER_ROW = 3
COL_TICKER = "A"
COL_K = "K"
COL_M = "M"
COL_P = "P"
COL_Q = "Q"
COL_T = "T"
COL_U = "U"
COL_V = "V"
COL_S = "S"
COL_Y = "Y"
COL_Z = "Z"

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
LIVE_TRADE_INFO_FILE = os.path.join(_BASE_DIR, "Live_Trade_Info.xlsx")


@dataclass
class EarningsRow:
    row: int
    ticker: str
    m_token: str
    trade_date: date


def _normalize_direction(cell_value: Any) -> Optional[str]:
    if cell_value is None:
        return None
    s = str(cell_value).strip().lower()
    return s if s in ("long", "short") else None


def _collect_earnings_rows(sheet) -> List[EarningsRow]:
    try:
        max_row = int(sheet.used_range.last_cell.row)
    except Exception:
        max_row = HEADER_ROW + 500
    start = HEADER_ROW + 1
    out: List[EarningsRow] = []
    for row in range(start, max_row + 1):
        m_token = parse_m_dow(sheet.range(f"{COL_M}{row}").value)
        k_date = parse_earnings_date(sheet.range(f"{COL_K}{row}").value)
        direction = _normalize_direction(sheet.range(f"{COL_Y}{row}").value)
        ticker_raw = sheet.range(f"{COL_TICKER}{row}").value
        if not m_token or k_date is None or not direction:
            continue
        if ticker_raw is None or str(ticker_raw).strip() == "":
            continue
        ticker = str(ticker_raw).strip().upper()
        trade_date = resolve_trade_date(m_token, k_date)
        out.append(
            EarningsRow(row=row, ticker=ticker, m_token=m_token, trade_date=trade_date)
        )
    return out


def _contiguous_groups(rows: Sequence[EarningsRow]) -> List[List[EarningsRow]]:
    """Split into contiguous row blocks that share the same M token."""
    if not rows:
        return []
    ordered = sorted(rows, key=lambda r: r.row)
    groups: List[List[EarningsRow]] = []
    cur = [ordered[0]]
    for r in ordered[1:]:
        prev = cur[-1]
        if r.m_token == prev.m_token and r.row == prev.row + 1:
            cur.append(r)
        else:
            groups.append(cur)
            cur = [r]
    groups.append(cur)
    return groups


def _first_empty_dest_index(sheet, group: Sequence[EarningsRow], dest_col: str) -> Optional[int]:
    for i, er in enumerate(group):
        if not cell_has_value(sheet.range(f"{dest_col}{er.row}").value):
            return i
    return None


def _is_stop_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    if isinstance(cell_value, (int, float)):
        return cell_value == 0
    return str(cell_value).strip() == "0"


def _is_q_c_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "C"


def _is_q_chain_flag(cell_value: Any) -> bool:
    """True for M2 / M1 / C (do not overwrite these with a stop 0)."""
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() in ("M2", "M1", "C")


def _sheet_max_row(sheet) -> int:
    try:
        return int(sheet.used_range.last_cell.row)
    except Exception:
        return HEADER_ROW + 500


def _sheet_scan_max_row(sheet) -> int:
    """
    Upper bound for flag scans: used_range plus M day flag, ticker, or zz end marker.
    Avoids missing rows just beyond Excel used_range.
    """
    max_row = _sheet_max_row(sheet)
    # Probe a cushion below used_range for M/ticker/zz that Excel has not marked used yet
    probe_to = max(max_row + 50, HEADER_ROW + 500)
    last = max_row
    for row in range(HEADER_ROW + 1, probe_to + 1):
        m = sheet.range(f"{COL_M}{row}").value
        a = sheet.range(f"{COL_TICKER}{row}").value
        has_m = parse_m_dow(m) is not None
        has_a = a is not None and str(a).strip() != ""
        if has_m or has_a:
            last = max(last, row)
    return last


def _find_next_c_row(sheet, after_row: int, max_row: int) -> Optional[int]:
    """First Q C at or below after_row (inclusive)."""
    for row in range(after_row, max_row + 1):
        if _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value):
            return row
    return None


def _write_q_zero_safe(sheet, zero_row: int) -> bool:
    """Write Q stop 0 only if it would not overwrite M2/M1/C or land on a trade row."""
    if _row_is_trade_or_ticker(sheet, zero_row):
        print(f"  Skip Q 0@{zero_row}: trade/ticker row (needs C, not 0).")
        return False
    existing = sheet.range(f"{COL_Q}{zero_row}").value
    if _is_q_chain_flag(existing):
        print(
            f"  Skip Q 0@{zero_row}: would overwrite existing "
            f"{str(existing).strip()!r}."
        )
        return False
    sheet.range(f"{COL_Q}{zero_row}").value = 0
    return True


def _clear_p_flag_span(sheet, start_row: int, zero_row: int) -> None:
    """Clear only column P between start and terminator (never touch Q)."""
    for r in range(start_row, zero_row + 1):
        sheet.range(f"{COL_P}{r}").value = None


def _write_o_range(sheet, start_row: int, zero_row: int) -> int:
    """
    Column P: O at start, 0 at zero_row.
    zero_row is often the same row as Q's C (P=0 and Q=C coexist).
    """
    _clear_p_flag_span(sheet, start_row, zero_row)
    sheet.range(f"{COL_P}{start_row}").value = "O"
    sheet.range(f"{COL_P}{zero_row}").value = 0
    return zero_row


def _write_q_start_only(sheet, flag: str, start_row: int) -> None:
    """Write M2 or M1 at start_row only — next flag in the Q chain ends this range."""
    sheet.range(f"{COL_Q}{start_row}").value = flag


def _clear_stale_q_stop_after_block(sheet, last_data_row: int) -> None:
    """
    After converting a block to M1/M2, remove a leftover Q 0 on the next row.
    That stale 0 is from a prior C-range end and often sits on the next block's C cell.
    Never clear an actual C/M1/M2 flag.
    """
    after = last_data_row + 1
    val = sheet.range(f"{COL_Q}{after}").value
    if _is_stop_flag(val) and not _is_q_chain_flag(val):
        sheet.range(f"{COL_Q}{after}").value = None
        print(f"  Cleared stale Q 0@{after} after M1/M2 block (protect C row).")


# Max rows below C to search for column-A "zz" (end of trade table) when placing Q 0.
C_RANGE_ZERO_MAX_OFFSET = 500


def _is_zz_end_marker(cell_value: Any) -> bool:
    """Column A 'zz' marks the end of the trade table (not a ticker)."""
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "ZZ"


def _row_is_trade_or_ticker(sheet, row: int) -> bool:
    """True if column M has a day flag OR column A has a real ticker (not zz)."""
    if parse_m_dow(sheet.range(f"{COL_M}{row}").value) is not None:
        return True
    a = sheet.range(f"{COL_TICKER}{row}").value
    if a is None or str(a).strip() == "":
        return False
    if _is_zz_end_marker(a):
        return False
    return True


def _find_zz_end_row(sheet, start_row: int) -> Optional[int]:
    """First column-A zz at or below start_row (end of trades marker)."""
    max_row = _sheet_scan_max_row(sheet)
    probe_to = max(max_row, start_row + C_RANGE_ZERO_MAX_OFFSET)
    for row in range(start_row, probe_to + 1):
        if _is_zz_end_marker(sheet.range(f"{COL_TICKER}{row}").value):
            return row
    return None


def _resolve_q_zero_row_after_c(sheet, c_row: int) -> int:
    """
    Row for the Q 0 that ends the C range: the column-A 'zz' end-of-trades row.
    Falls back only if zz is missing (should not happen on a normal sheet).
    """
    zz_row = _find_zz_end_row(sheet, c_row + 1)
    if zz_row is not None:
        return zz_row
    # Fallback: first non-trade row below C (legacy)
    print(f"  Warning: no column-A zz below C@{c_row}; falling back to first blank trade row.")
    for offset in range(1, C_RANGE_ZERO_MAX_OFFSET + 1):
        row = c_row + offset
        if not _row_is_trade_or_ticker(sheet, row):
            return row
    return c_row + C_RANGE_ZERO_MAX_OFFSET


def _clear_stale_q_zeros_between(sheet, start_row_exclusive: int, end_row_exclusive: int) -> None:
    """Remove leftover Q stop 0s in (start, end) so Get_Closes does not stop early."""
    for r in range(start_row_exclusive + 1, end_row_exclusive):
        val = sheet.range(f"{COL_Q}{r}").value
        if _is_stop_flag(val) and not _is_q_chain_flag(val):
            sheet.range(f"{COL_Q}{r}").value = None
            print(f"  Cleared stale Q 0@{r} before new C-range terminator.")


def _clear_stale_q_zeros_after(sheet, after_row: int) -> None:
    """Remove Q stop 0s below the zz terminator (e.g. prior far placements)."""
    max_row = _sheet_scan_max_row(sheet)
    for r in range(after_row + 1, max_row + 1):
        val = sheet.range(f"{COL_Q}{r}").value
        if _is_stop_flag(val) and not _is_q_chain_flag(val):
            sheet.range(f"{COL_Q}{r}").value = None
            print(f"  Cleared stale Q 0@{r} below zz terminator.")


def _place_q_zero_after_c(sheet, c_row: int) -> int:
    """
    Write the Q 0 that closes the C range on the column-A zz row.
    Clears stale early 0s between C and zz, and any 0s below zz.
    """
    zero_row = _resolve_q_zero_row_after_c(sheet, c_row)
    _clear_stale_q_zeros_between(sheet, c_row, zero_row)
    _clear_stale_q_zeros_after(sheet, zero_row)
    if _write_q_zero_safe(sheet, zero_row):
        zz_note = " (zz)" if _is_zz_end_marker(sheet.range(f"{COL_TICKER}{zero_row}").value) else ""
        print(f"  Q C-range terminator 0@{zero_row}{zz_note} (C@{c_row})")
    return zero_row


def _write_q_c_range(sheet, start_row: int, last_data_row: int) -> int:
    """Column Q: C at start; 0 on the column-A zz end-of-trades row."""
    for r in range(start_row, last_data_row + 1):
        sheet.range(f"{COL_Q}{r}").value = None
    sheet.range(f"{COL_Q}{start_row}").value = "C"
    return _place_q_zero_after_c(sheet, start_row)


def _ensure_c_row_for_m1_o(
    sheet,
    *,
    m1_start_row: int,
    last_data_row: int,
    max_row: int,
) -> int:
    """
    Ensure a Q C exists below the M1 block for chain end + P closing 0.

    - Prefer an existing C below the M1 data rows.
    - If none, write C at last_data_row+1 (never leave that cell as Q 0).
    - Place Q stop 0 on the column-A zz end-of-trades row after C.
    Returns the C row (same row as P's closing 0).
    """
    c_row = _find_next_c_row(sheet, last_data_row + 1, max_row)
    if c_row is None:
        c_row = _find_next_c_row(sheet, m1_start_row + 1, max_row)
    if c_row is not None and c_row <= m1_start_row:
        c_row = None

    created = False
    if c_row is None:
        c_row = last_data_row + 1
        created = True

    q_here = sheet.range(f"{COL_Q}{c_row}").value
    # Stale Q 0 on the C row (classic bug) or missing C -> write C.
    if created or _is_stop_flag(q_here) or not _is_q_c_flag(q_here):
        if _is_q_m1_flag(q_here) or _is_q_m2_flag(q_here):
            # Should not land C on M1/M2; push one row down.
            c_row += 1
            q_here = sheet.range(f"{COL_Q}{c_row}").value
        sheet.range(f"{COL_Q}{c_row}").value = "C"
        print(
            f"  Ensured Q C@{c_row} for M1 end / P closing-0 row"
            + (" (created)" if created else " (replaced stale Q 0/empty)")
        )

    _place_q_zero_after_c(sheet, c_row)
    return c_row



def _is_o_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "O"


def _is_q_m1_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "M1"


def _is_q_m2_flag(cell_value: Any) -> bool:
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() == "M2"


class FlagLayoutError(ValueError):
    """Raised when post-flag QC invariants fail."""


def _scan_flag_rows(sheet) -> dict:
    """Collect P/Q flag row numbers from the Trades data area."""
    max_row = _sheet_scan_max_row(sheet)
    start = HEADER_ROW + 1
    out = {
        "o": [],
        "p_zero": [],
        "m2": [],
        "m1": [],
        "c": [],
        "q_zero": [],
    }
    for row in range(start, max_row + 1):
        p = sheet.range(f"{COL_P}{row}").value
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_o_flag(p):
            out["o"].append(row)
        if _is_stop_flag(p):
            out["p_zero"].append(row)
        if _is_q_m2_flag(q):
            out["m2"].append(row)
        if _is_q_m1_flag(q):
            out["m1"].append(row)
        if _is_q_c_flag(q):
            out["c"].append(row)
        if _is_stop_flag(q):
            out["q_zero"].append(row)
    return out



def _is_trade_row(sheet, row: int) -> bool:
    """True if column M has a day flag OR column A has a ticker."""
    return _row_is_trade_or_ticker(sheet, row)


def _trade_row_c_rows(sheet) -> List[int]:
    max_row = _sheet_scan_max_row(sheet)
    out: List[int] = []
    for row in range(HEADER_ROW + 1, max_row + 1):
        if _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value) and _is_trade_row(sheet, row):
            out.append(row)
    return out


def _last_weekday_block_start_row(sheet) -> Optional[int]:
    """
    Start row of the last contiguous column-M weekday block above zz.
    Contiguous = consecutive rows sharing the same M token with ticker present.
    """
    zz = _find_zz_end_row(sheet, HEADER_ROW + 1)
    end = (zz - 1) if zz is not None else _sheet_scan_max_row(sheet)
    last_start: Optional[int] = None
    prev_row: Optional[int] = None
    prev_m: Optional[str] = None
    for row in range(HEADER_ROW + 1, end + 1):
        m_token = parse_m_dow(sheet.range(f"{COL_M}{row}").value)
        ticker = _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{row}").value)
        if m_token is None or ticker is None:
            continue
        if prev_row is None or row != prev_row + 1 or m_token != prev_m:
            last_start = row
        prev_row = row
        prev_m = m_token
    return last_start


def _last_weekday_block_is_m1_or_m2(sheet) -> bool:
    """
    Only case where C is not required: the sheet's last M weekday block
    (just above zz) already starts with M1 or M2.
    """
    start = _last_weekday_block_start_row(sheet)
    if start is None:
        return False
    q = sheet.range(f"{COL_Q}{start}").value
    return _is_q_m1_flag(q) or _is_q_m2_flag(q)


def _first_unentered_trade_row(sheet) -> Optional[int]:
    """
    First trade row above zz with an M day flag, a real ticker, and blank column T.
    Blank T means no open yet — trade not entered — C belongs here.
    Skips rows that already hold M1/M2 (do not overwrite those flags).
    """
    zz = _find_zz_end_row(sheet, HEADER_ROW + 1)
    end = (zz - 1) if zz is not None else _sheet_scan_max_row(sheet)
    for row in range(HEADER_ROW + 1, end + 1):
        if parse_m_dow(sheet.range(f"{COL_M}{row}").value) is None:
            continue
        if not _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{row}").value):
            continue
        if cell_has_value(sheet.range(f"{COL_T}{row}").value):
            continue
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_q_m1_flag(q) or _is_q_m2_flag(q):
            continue
        return row
    return None


def _clear_trade_row_cs_except(sheet, keep_row: Optional[int] = None) -> None:
    max_row = _sheet_scan_max_row(sheet)
    for row in range(HEADER_ROW + 1, max_row + 1):
        if keep_row is not None and row == keep_row:
            continue
        if _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value) and _is_trade_row(sheet, row):
            sheet.range(f"{COL_Q}{row}").value = None
            print(f"  Cleared misplaced Q C@{row}")


def _first_trade_row_after(sheet, after_row: int) -> Optional[int]:
    """First M+ticker trade row strictly below after_row and above zz (not M1/M2)."""
    zz = _find_zz_end_row(sheet, HEADER_ROW + 1)
    end = (zz - 1) if zz is not None else _sheet_scan_max_row(sheet)
    for row in range(after_row + 1, end + 1):
        if parse_m_dow(sheet.range(f"{COL_M}{row}").value) is None:
            continue
        if not _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{row}").value):
            continue
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_q_m1_flag(q) or _is_q_m2_flag(q):
            continue
        return row
    return None


def ensure_required_c_flag(
    sheet,
    *,
    after_open_block_row: Optional[int] = None,
) -> None:
    """
    Always keep a trade-row C, except when the last M weekday block above zz
    is already M1 or M2.

    Placement:
      - Evening after open-trades: first trade row after the entered weekday block
      - Otherwise: first unentered trade (blank column T)

    Closing Q 0 stays on the column-A zz row.
    """
    print("\nEnsuring required C...")

    max_row = _sheet_scan_max_row(sheet)
    for row in range(HEADER_ROW + 1, max_row + 1):
        if _is_stop_flag(sheet.range(f"{COL_Q}{row}").value) and _row_is_trade_or_ticker(
            sheet, row
        ):
            sheet.range(f"{COL_Q}{row}").value = None
            print(f"  Cleared Q 0@{row} on trade/ticker row")

    if _last_weekday_block_is_m1_or_m2(sheet):
        start = _last_weekday_block_start_row(sheet)
        print(
            f"  Last M weekday block @{start} is M1/M2; C not required."
        )
        _clear_trade_row_cs_except(sheet, keep_row=None)
        return

    if after_open_block_row is not None:
        c_at = _first_trade_row_after(sheet, after_open_block_row)
        reason = f"first trade after open block ending @{after_open_block_row}"
    else:
        c_at = _first_unentered_trade_row(sheet)
        reason = "first unentered trade: blank T"

    if c_at is None:
        print(f"  No row found for required C ({reason}).")
        return

    _clear_trade_row_cs_except(sheet, keep_row=c_at)

    if not _is_q_c_flag(sheet.range(f"{COL_Q}{c_at}").value):
        sheet.range(f"{COL_Q}{c_at}").value = "C"
        print(f"  Placed Q C@{c_at} ({reason})")
    else:
        print(f"  Q C already at @{c_at} ({reason})")

    _place_q_zero_after_c(sheet, c_at)


def _m1_is_terminal_trade_block(sheet) -> bool:
    """ Backward-compatible name: last weekday block is M1 or M2. """
    return _last_weekday_block_is_m1_or_m2(sheet)


def _ensure_c_after_m2_if_needed(
    sheet,
    last_m2_data_row: int,
    all_rows: Sequence[EarningsRow] = (),
    as_of: Optional[date] = None,
) -> bool:
    """After M2 writes, defer to blank-T C rule."""
    del last_m2_data_row, all_rows, as_of
    before = _trade_row_c_rows(sheet)
    ensure_required_c_flag(sheet)
    return bool(_trade_row_c_rows(sheet)) and (
        not before or _trade_row_c_rows(sheet) != before
    )


def repair_q_layout_before_qc(
    sheet,
    wrote_m2_last: Optional[List[int]] = None,
    all_rows: Optional[Sequence[EarningsRow]] = None,
    as_of: Optional[date] = None,
) -> None:
    """Final C layout: first blank-T trade, unless last block is M1/M2."""
    del wrote_m2_last, all_rows, as_of
    ensure_required_c_flag(sheet)


def _is_evening_flag_mode(sheet, calendar_today: date) -> bool:
    both_ok, _, _ = opens_and_closes_succeeded_on(calendar_today)
    return both_ok and _active_flag_dests_complete(sheet)


def _evening_open_orders_ready(calendar_today: date) -> bool:
    """
    Evening-only gate for promoting today's entries to M1+O:
      both Open_Trades_ToS and Open_Trade_ToS2 succeeded today with trades,
      and Live_Trade_Info still has trades on today's date sheet.
    """
    both_ok, d1, d2 = open_trades_both_succeeded_on(calendar_today)
    live_ok = live_info_has_trades_for_day(LIVE_TRADE_INFO_FILE, calendar_today)
    print("\nEvening open-orders QC:")
    print(f"  Open_Trades_ToS success date (MT): {d1}")
    print(f"  Open_Trade_ToS2 success date (MT): {d2}")
    print(f"  Both open scripts succeeded with trades today: {both_ok}")
    print(f"  Live_Trade_Info has trades for {calendar_today}: {live_ok}")
    return both_ok and live_ok


def place_evening_entered_trade_flags(
    sheet,
    all_rows: Sequence[EarningsRow],
    calendar_today: date,
) -> Optional[int]:
    """
    After market close (evening mode + open orders today):
      M1 + O on calendar-today's weekday earnings block(s).
    Returns the last row of the last such block (C goes on the next trade row).
    """
    token = dow_abbrev_for_date(calendar_today)
    groups = [
        g
        for g in _contiguous_groups(all_rows)
        if g[0].trade_date == calendar_today and g[0].m_token == token
    ]
    if not groups:
        groups = [
            g for g in _contiguous_groups(all_rows) if g[0].m_token == token
        ]
    if not groups:
        print(
            f"  Evening open-trades: no earnings groups with M={token} "
            f"for {calendar_today}; skip M1/O."
        )
        return None

    print(
        f"\nEvening open-trades: placing M1+O for entered block(s) "
        f"M={token} D={calendar_today}..."
    )
    max_row = _sheet_max_row(sheet)
    last_end: Optional[int] = None
    for g in groups:
        start_row = g[0].row
        last_data_row = g[-1].row
        tickers = ", ".join(r.ticker for r in g)
        # Force M1+O even if T already has values (orders already live).
        if _is_q_c_flag(sheet.range(f"{COL_Q}{start_row}").value):
            sheet.range(f"{COL_Q}{start_row}").value = None
        c_row = _ensure_c_row_for_m1_o(
            sheet,
            m1_start_row=start_row,
            last_data_row=last_data_row,
            max_row=max_row,
        )
        _write_q_start_only(sheet, "M1", start_row)
        z_p = _write_o_range(sheet, start_row, c_row)
        print(
            f"  Evening O@{start_row} 0@{z_p}; Q M1@{start_row} ends at C@{c_row}; "
            f"tickers: {tickers}"
        )
        last_end = last_data_row
    return last_end


def lock_share_sizes_for_next_trading_day(
    sheet,
    all_rows: Sequence[EarningsRow],
    calendar_today: date,
) -> None:
    """
    Evening-only: for next trading day's planned trades (Live_Trade_Info guide),
    paste column Z as values so share size is locked (no formula).

    If Live_Trade_Info has no tomorrow sheet/trades yet, write them from
    earnings rows with trade_date == tomorrow (using evaluated Z), then lock.
    """
    tomorrow = next_trading_day(calendar_today)
    print(f"\nLocking Z share sizes for next trading day {tomorrow}...")

    tomorrow_earnings = [r for r in all_rows if r.trade_date == tomorrow]
    if not tomorrow_earnings:
        print(f"  No earnings rows with trade_date={tomorrow}; skip Z lock.")
        return

    live_pairs = read_live_trades_for_day(LIVE_TRADE_INFO_FILE, tomorrow)
    live_keys = {(t, d) for t, d in live_pairs}

    if not live_keys:
        print(
            f"  Live_Trade_Info has no trades for {tomorrow}; "
            f"writing {len(tomorrow_earnings)} trade(s) from earnings, then locking Z."
        )
        # Evaluate Z first, then write Live_Trade_Info with numeric sizes
        to_write = []
        for r in tomorrow_earnings:
            direction = _normalize_direction(sheet.range(f"{COL_Y}{r.row}").value)
            if direction not in ("long", "short"):
                print(f"  Skip row {r.row} {r.ticker}: bad direction")
                continue
            z_cell = sheet.range(f"{COL_Z}{r.row}")
            val = z_cell.value
            try:
                size_num = int(round(float(val)))
            except (TypeError, ValueError):
                print(f"  Skip row {r.row} {r.ticker}: Z not numeric ({val!r})")
                continue
            to_write.append((r.ticker, direction, size_num))
            live_keys.add((r.ticker, direction))
        if to_write:
            try:
                sheet_name = write_live_trades_for_day(
                    LIVE_TRADE_INFO_FILE, tomorrow, to_write
                )
                print(f"  Wrote {len(to_write)} trade(s) to Live_Trade_Info '{sheet_name}'.")
            except Exception as e:
                print(f"  Warning: could not write Live_Trade_Info for {tomorrow}: {e}")
        else:
            print("  Nothing to write/lock for tomorrow.")
            return
    else:
        print(
            f"  Live_Trade_Info guide: {len(live_keys)} trade(s) for {tomorrow}; "
            f"locking matching earnings Z cells."
        )

    locked = 0
    for r in tomorrow_earnings:
        direction = _normalize_direction(sheet.range(f"{COL_Y}{r.row}").value)
        if (r.ticker, direction) not in live_keys:
            continue
        z_cell = sheet.range(f"{COL_Z}{r.row}")
        try:
            formula = z_cell.formula
        except Exception:
            formula = None
        val = z_cell.value
        if val is None or str(val).strip() == "":
            print(f"  Skip Z{r.row} {r.ticker}: empty")
            continue
        try:
            size_num = int(round(float(val)))
        except (TypeError, ValueError):
            print(f"  Skip Z{r.row} {r.ticker}: cannot coerce {val!r}")
            continue

        has_formula = isinstance(formula, str) and formula.startswith("=")
        if has_formula or z_cell.value != size_num:
            z_cell.value = size_num
            print(f"  Locked Z{r.row} {r.ticker} ({direction}) = {size_num}")
            locked += 1
        else:
            print(f"  Z{r.row} {r.ticker} already locked value {size_num}")

    print(f"  Z lock complete: {locked} cell(s) pasted as values.")


def validate_flag_layout_strict(sheet) -> None:
    """
    Strict post-flag QC. Raises FlagLayoutError if invariants fail.

    Rules (when the relevant flags are present):
      1. O (P) and M1 (Q) are on the exact same row(s)
      2. The P 0 that closes O is on the same row(s) as C (Q)
      3. M2 row number(s) are strictly less than M1 row number(s)
      4. If M2 exists with C but no M1, M2 row(s) < C row(s)
    """
    flags = _scan_flag_rows(sheet)
    o_rows = flags["o"]
    p_zero = flags["p_zero"]
    m2_rows = flags["m2"]
    m1_rows = flags["m1"]
    c_rows = flags["c"]
    errors: List[str] = []

    print("\nFlag layout QC (strict):")
    print(f"  P O@{o_rows or '-'}  0@{p_zero or '-'}")
    print(
        f"  Q M2@{m2_rows or '-'}  M1@{m1_rows or '-'}  "
        f"C@{c_rows or '-'}  0@{flags['q_zero'] or '-'}"
    )

    # 1) O and M1 same rows
    if o_rows or m1_rows:
        if set(o_rows) != set(m1_rows):
            errors.append(
                f"O and M1 must be on the same row(s); O@{o_rows} M1@{m1_rows}"
            )

    # 2) P closing 0 aligned with C whenever O range exists, or both P0 and C exist
    if o_rows:
        if not c_rows:
            errors.append(
                f"O is present at {o_rows} but no C flag in Q "
                f"(P closing 0 must share C row)"
            )
        if not p_zero:
            errors.append(
                f"O is present at {o_rows} but no closing 0 in column P"
            )
        if c_rows and p_zero and set(p_zero) != set(c_rows):
            errors.append(
                f"P closing 0 must be on the same row(s) as C; "
                f"P 0@{p_zero} C@{c_rows}"
            )
    elif p_zero or c_rows:
        if p_zero and c_rows and set(p_zero) != set(c_rows):
            errors.append(
                f"P closing 0 must be on the same row(s) as C; "
                f"P 0@{p_zero} C@{c_rows}"
            )
        if p_zero and not c_rows:
            errors.append(f"P closing 0 at {p_zero} but no C flag in Q")

    # 3) M2 above M1
    if m2_rows and m1_rows:
        if max(m2_rows) >= min(m1_rows):
            errors.append(
                f"M2 row(s) must be less than M1 row(s); "
                f"M2@{m2_rows} M1@{m1_rows}"
            )

    # 4) M2 above C when no M1
    if m2_rows and c_rows and not m1_rows:
        if max(m2_rows) >= min(c_rows):
            errors.append(
                f"M2 row(s) must be less than C row(s) when no M1; "
                f"M2@{m2_rows} C@{c_rows}"
            )

    # 5) Q 0 must not sit on a trade row (that cell should be C)
    max_row = _sheet_scan_max_row(sheet)
    for z in flags["q_zero"]:
        if _is_trade_row(sheet, z):
            errors.append(
                f"Q 0@{z} is on a trade/ticker row; "
                f"expected C there for future closes/sizing"
            )

    # 6) Always require a trade-row C, unless last M weekday block is M1/M2
    trade_cs = [r for r in c_rows if _is_trade_row(sheet, r)]
    if not trade_cs and not _last_weekday_block_is_m1_or_m2(sheet):
        errors.append(
            "Missing trade-row C in column Q "
            "(required unless last M weekday block above zz is M1 or M2)"
        )

    if errors:
        for msg in errors:
            print(f"  QC FAIL: {msg}")
        raise FlagLayoutError(
            "Flag layout QC failed:\n- " + "\n- ".join(errors)
        )

    print("  QC PASS")


def _rows_for_trade_date(all_rows: Sequence[EarningsRow], d: date) -> List[EarningsRow]:
    return [r for r in all_rows if r.trade_date == d]


def _flag_job_for_groups(
    sheet,
    groups: List[List[EarningsRow]],
    *,
    dest_col: str,
    q_flag: Optional[str] = None,
    p_flag: Optional[str] = None,
    job_label: str,
) -> List[int]:
    """
    For each contiguous DOW group, find first empty dest cell and write flags.

    Q chain: M2 until M1; M1 until C; C until 0.
    O (P) trailing 0 may share the C row; never write a Q 0 for M1/M2.
    """
    wrote_last_rows: List[int] = []
    wrote_any = False
    max_row = _sheet_max_row(sheet)

    for group in groups:
        idx = _first_empty_dest_index(sheet, group, dest_col)
        if idx is None:
            print(
                f"  {job_label}: group M={group[0].m_token} rows "
                f"{group[0].row}-{group[-1].row} already filled in {dest_col}; skip."
            )
            continue
        sub = group[idx:]
        start_row = sub[0].row
        last_data_row = sub[-1].row
        tickers = ", ".join(r.ticker for r in sub)

        if p_flag and q_flag:
            # O + M1: discover T and U starts independently; must match.
            idx_u = _first_empty_dest_index(sheet, group, COL_U)
            idx_t = _first_empty_dest_index(sheet, group, COL_T)
            if idx_u is None and idx_t is None:
                print(f"  {job_label}: T and U already filled for M={group[0].m_token}; skip.")
                continue
            if idx_u != idx_t:
                print(
                    f"  ERROR {job_label}: O start (T empty idx={idx_t}) != "
                    f"M1 start (U empty idx={idx_u}) for M={group[0].m_token}; "
                    f"writing neither."
                )
                continue
            assert idx_t is not None
            sub = group[idx_t:]
            start_row = sub[0].row
            last_data_row = sub[-1].row
            tickers = ", ".join(r.ticker for r in sub)

            # M1 ends at C; P closing 0 shares the C row. Never leave Q 0 on C.
            c_row = _ensure_c_row_for_m1_o(
                sheet,
                m1_start_row=start_row,
                last_data_row=last_data_row,
                max_row=max_row,
            )
            if last_data_row >= c_row:
                last_data_row = c_row - 1
            p_zero_row = c_row

            _write_q_start_only(sheet, "M1", start_row)
            # Do not clear Q on the C row (that is p_zero_row).
            if last_data_row + 1 < c_row:
                _clear_stale_q_stop_after_block(sheet, last_data_row)
            z_p = _write_o_range(sheet, start_row, p_zero_row)
            print(
                f"  {job_label}: P O@{start_row} 0@{z_p}; "
                f"Q M1@{start_row} ends at C@{c_row}; tickers: {tickers}"
            )
            wrote_last_rows.append(last_data_row)
            wrote_any = True
            continue

        if q_flag == "C":
            z_q = _write_q_c_range(sheet, start_row, last_data_row)
            print(
                f"  {job_label}: Q C@{start_row} 0@{z_q} "
                f"(M={group[0].m_token}); tickers: {tickers}"
            )
            wrote_last_rows.append(last_data_row)
            wrote_any = True
            continue

        if q_flag == "M2":
            _write_q_start_only(sheet, "M2", start_row)
            _clear_stale_q_stop_after_block(sheet, last_data_row)
            print(
                f"  {job_label}: Q M2@{start_row} "
                f"(ends at M1/C, no Q 0); (M={group[0].m_token}); tickers: {tickers}"
            )
            wrote_last_rows.append(last_data_row)
            wrote_any = True

    if not wrote_any:
        print(f"  No Full_Auto rows to flag for {job_label}.")
    return wrote_last_rows


def _normalize_ticker_cell(cell_value: Any) -> Optional[str]:
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper()
    if not t or t == "ZZ":
        return None
    return t


def _dest_range_complete(sheet, start_row: int, dest_col: str, max_row: int) -> bool:
    """True if every ticker row from start until next Q chain flag/0 has dest filled."""
    saw = False
    for row in range(start_row, max_row + 1):
        if row > start_row:
            q = sheet.range(f"{COL_Q}{row}").value
            if _is_stop_flag(q) or _is_q_chain_flag(q):
                break
        if not _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{row}").value):
            continue
        saw = True
        if not cell_has_value(sheet.range(f"{dest_col}{row}").value):
            return False
    return saw


def _o_range_t_complete(sheet, start_row: int, max_row: int) -> bool:
    """True if every ticker from O until P 0 has T filled (or no O on start row)."""
    if not _is_o_flag(sheet.range(f"{COL_P}{start_row}").value):
        return True
    for row in range(start_row, max_row + 1):
        p = sheet.range(f"{COL_P}{row}").value
        if row > start_row and _is_stop_flag(p):
            break
        if not _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{row}").value):
            continue
        if not cell_has_value(sheet.range(f"{COL_T}{row}").value):
            return False
    return True


def _c_range_has_ticker(sheet, start_row: int, max_row: int) -> bool:
    for r2 in range(start_row, max_row + 1):
        if r2 > start_row:
            q2 = sheet.range(f"{COL_Q}{r2}").value
            if _is_stop_flag(q2) or _is_q_chain_flag(q2):
                break
        if _normalize_ticker_cell(sheet.range(f"{COL_TICKER}{r2}").value):
            return True
    return False


def _active_flag_dests_complete(sheet) -> bool:
    """QC: every active M2/M1/C data range has dest prices filled."""
    max_row = _sheet_max_row(sheet)
    found = False
    for row in range(HEADER_ROW + 1, max_row + 1):
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_q_m2_flag(q):
            found = True
            if not _dest_range_complete(sheet, row, COL_V, max_row):
                print(f"  QC prices: M2@{row} still has empty V")
                return False
        elif _is_q_m1_flag(q):
            found = True
            if not _dest_range_complete(sheet, row, COL_U, max_row):
                print(f"  QC prices: M1@{row} still has empty U")
                return False
            if not _o_range_t_complete(sheet, row, max_row):
                print(f"  QC prices: O/M1@{row} still has empty T")
                return False
        elif _is_q_c_flag(q):
            if not _c_range_has_ticker(sheet, row, max_row):
                continue
            found = True
            if not _dest_range_complete(sheet, row, COL_S, max_row):
                print(f"  QC prices: C@{row} still has empty S")
                return False
    if found:
        print("  QC prices: active flag dest columns are filled")
    else:
        print("  QC prices: no active M2/M1/C data ranges (OK)")
    return True


def _clear_p_o_range(sheet, start_row: int, max_row: int) -> None:
    sheet.range(f"{COL_P}{start_row}").value = None
    for row in range(start_row, max_row + 1):
        if _is_stop_flag(sheet.range(f"{COL_P}{row}").value):
            sheet.range(f"{COL_P}{row}").value = None
            print(f"  Cleared P O-range stop 0@{row}")
            break


def _clear_q_zeros_after(sheet, start_row: int, max_row: int) -> None:
    for row in range(start_row + 1, min(max_row, start_row + C_RANGE_ZERO_MAX_OFFSET) + 1):
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_q_chain_flag(q):
            break
        if _is_stop_flag(q):
            sheet.range(f"{COL_Q}{row}").value = None
            print(f"  Cleared Q stop 0@{row} after completed C")


def clear_completed_price_flags(sheet) -> None:
    """Remove M2/M1/O/C flags whose destination price columns are fully populated."""
    max_row = _sheet_max_row(sheet)
    m2_rows: List[int] = []
    m1_rows: List[int] = []
    c_rows: List[int] = []
    for row in range(HEADER_ROW + 1, max_row + 1):
        q = sheet.range(f"{COL_Q}{row}").value
        if _is_q_m2_flag(q):
            m2_rows.append(row)
        elif _is_q_m1_flag(q):
            m1_rows.append(row)
        elif _is_q_c_flag(q):
            c_rows.append(row)

    print("\nClearing completed price flags (dest columns full)...")
    for row in m2_rows:
        if _dest_range_complete(sheet, row, COL_V, max_row):
            sheet.range(f"{COL_Q}{row}").value = None
            print(f"  Cleared M2@{row} (V complete)")

    for row in m1_rows:
        if _dest_range_complete(sheet, row, COL_U, max_row) and _o_range_t_complete(
            sheet, row, max_row
        ):
            sheet.range(f"{COL_Q}{row}").value = None
            print(f"  Cleared M1@{row} (U/T complete)")
            _clear_p_o_range(sheet, row, max_row)

    for row in c_rows:
        if not _c_range_has_ticker(sheet, row, max_row):
            sheet.range(f"{COL_Q}{row}").value = None
            _clear_q_zeros_after(sheet, row, max_row)
            print(f"  Cleared terminator C@{row}")
            continue
        if _dest_range_complete(sheet, row, COL_S, max_row):
            sheet.range(f"{COL_Q}{row}").value = None
            _clear_q_zeros_after(sheet, row, max_row)
            print(f"  Cleared C@{row} (S complete)")


def resolve_flag_as_of_date(sheet, calendar_today: date) -> date:
    """
    Evening (opens+closes succeeded today AND dests filled):
      as-of = next trading day.
    Otherwise: as-of = calendar_today.
    """
    both_ok, opens_d, closes_d = opens_and_closes_succeeded_on(calendar_today)
    prices_ok = _active_flag_dests_complete(sheet)
    print("\nFlag as-of QC:")
    print(f"  Calendar today: {calendar_today}")
    print(f"  Get_Opens success date (MT): {opens_d}")
    print(f"  Get_Closes success date (MT): {closes_d}")
    print(f"  Opens+Closes both on today: {both_ok}")
    print(f"  Active flag dests complete: {prices_ok}")

    if both_ok and prices_ok:
        as_of = next_trading_day(calendar_today)
        print(f"  Evening mode -> place flags as-of next trading day: {as_of}")
        return as_of

    print(f"  Same-day/morning mode -> place flags as-of: {calendar_today}")
    return calendar_today

def flag_prices_full_auto(
    *,
    earnings_file: str = LATEST_EARNINGS_FILE,
    today: Optional[date] = None,
) -> None:
    """
    Write P/Q flags for Full_Auto price jobs.

    Evening (opens+closes succeeded today and dests filled): clear completed flags,
    then place flags as-of the *next* trading day.
    Morning / otherwise: place flags as-of calendar today (idempotent).
    """
    calendar_today = today or date.today()
    if not is_trading_day(calendar_today):
        print(f"Flag_Prices_Full_Auto: {calendar_today} is not a trading day; skip.")
        return
    if not os.path.exists(earnings_file):
        print(f"Flag_Prices_Full_Auto: earnings file not found: {earnings_file}")
        return

    print(f"\nFlag_Prices_Full_Auto (calendar {calendar_today})...")
    app = None
    wb = None
    owned_app = False
    owned_book = False
    try:
        app, wb, owned_app, owned_book = open_or_attach_earnings_workbook(earnings_file)
        try:
            sheet = wb.sheets[LATEST_EARNINGS_SHEET]
        except Exception:
            print(f"Sheet '{LATEST_EARNINGS_SHEET}' not found.")
            return

        all_rows = _collect_earnings_rows(sheet)
        if not all_rows:
            print("  No Full_Auto earnings rows (M+K+Y); nothing to flag.")
            return

        as_of = resolve_flag_as_of_date(sheet, calendar_today)
        evening = _is_evening_flag_mode(sheet, calendar_today)
        clear_completed_price_flags(sheet)

        open_block_end: Optional[int] = None
        if evening and _evening_open_orders_ready(calendar_today):
            open_block_end = place_evening_entered_trade_flags(
                sheet, all_rows, calendar_today
            )
        elif evening:
            print(
                "  Evening mode but open-orders gate not met; "
                "skipping entered-day M1/O promotion."
            )

        trade_dates = sorted({r.trade_date for r in all_rows})

        d_for_c = [d for d in trade_dates if d == as_of]
        d_for_m1 = [d for d in trade_dates if add_trading_days(d, 1) == as_of]
        d_for_m2 = [d for d in trade_dates if add_trading_days(d, 2) == as_of]

        print(
            f"\nPlacing flags as-of {as_of}: "
            f"C days={d_for_c}, M1 days={d_for_m1}, M2 days={d_for_m2}"
        )

        if not d_for_c and not d_for_m1 and not d_for_m2:
            print(
                f"  No date-matched C/M1/O/M2 jobs as-of {as_of}; "
                f"still ensuring required blank-T C."
            )
        else:
            c_start_rows: List[int] = []

            for d in d_for_c:
                rows = _rows_for_trade_date(all_rows, d)
                groups = _contiguous_groups(rows)
                for g in groups:
                    idx = _first_empty_dest_index(sheet, g, COL_S)
                    if idx is not None:
                        c_start_rows.append(g[idx].row)
                _flag_job_for_groups(
                    sheet, groups, dest_col=COL_S, q_flag="C", job_label=f"C->S (D={d})"
                )

            for d in d_for_m2:
                rows = _rows_for_trade_date(all_rows, d)
                groups = _contiguous_groups(rows)
                _flag_job_for_groups(
                    sheet, groups, dest_col=COL_V, q_flag="M2", job_label=f"M2->V (D={d})"
                )

            for d in d_for_m1:
                rows = _rows_for_trade_date(all_rows, d)
                groups = _contiguous_groups(rows)
                _flag_job_for_groups(
                    sheet,
                    groups,
                    dest_col=COL_T,
                    p_flag="O",
                    q_flag="M1",
                    job_label=f"O->T + M1->U (D={d}, D+1={as_of})",
                )

            for row in sorted(set(c_start_rows)):
                if not _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value):
                    sheet.range(f"{COL_Q}{row}").value = "C"
                    print(f"  Restored Q C@{row} after M1/M2 pass.")

        ensure_required_c_flag(sheet, after_open_block_row=open_block_end)

        if evening:
            lock_share_sizes_for_next_trading_day(sheet, all_rows, calendar_today)
        else:
            print(
                "\nSkipping Z share-size lock (morning/same-day mode; "
                "lock only runs after market close / evening mode)."
            )

        validate_flag_layout_strict(sheet)

        try:
            wb.save()
        except Exception:
            try:
                wb.save()
            except Exception as e:
                print(f"  Warning: could not save earnings workbook after flagging: {e}")
        print(f"Flag_Prices_Full_Auto: done (as-of {as_of}).")
    finally:
        release_earnings_workbook(app, wb, owned_app=owned_app, owned_book=owned_book)
