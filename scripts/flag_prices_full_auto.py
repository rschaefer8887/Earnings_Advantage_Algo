"""
Full_Auto price flagging for Latest Earnings columns P and Q.

Invoked from Stage_Trades_Auto when Trade_Mode is Full_Auto (same Task Scheduler job).

Price calendar (entry day D from M+K, trading days only):
  - today == D   -> Q: C at start; only Q terminator is 0 after the C group (closes -> S)
  - today == D+1 -> P: O ... 0 and Q: M1 at same start row (T and U)
  - today == D+2 -> Q: M2 at start (closes -> V)

Get_Closes_ToS Q chain: M2 until M1; M1 until C; C until 0.
  - M1 ends at C (never write a Q 0 for M1 — that 0 would land on the C row).
  - M2 ends at M1, or at C if no M1.
  - Only C gets a Q trailing 0 (placed at first blank column-M row within 100 rows below C, else C+100).
  - O (column P) trailing 0 must sit on the same row as C (P=0, Q=C).
  - If M1/O runs with no C below, create C on the terminator row and Q 0 after it.

No-op when there are no eligible rows for a job.

After writing flags, validate_flag_layout_strict() enforces:
  O row == M1 row; P closing 0 row == C row; M2 row < M1 row (or < C if no M1).
Failure raises FlagLayoutError (Stage exits non-zero).
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from typing import Any, List, Optional, Sequence

from full_auto_utils import (
    add_trading_days,
    cell_has_value,
    is_trading_day,
    parse_earnings_date,
    parse_m_dow,
    resolve_trade_date,
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


def _find_next_c_row(sheet, after_row: int, max_row: int) -> Optional[int]:
    """First Q C at or below after_row (inclusive)."""
    for row in range(after_row, max_row + 1):
        if _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value):
            return row
    return None


def _write_q_zero_safe(sheet, zero_row: int) -> bool:
    """Write Q stop 0 only if it would not overwrite M2/M1/C. Returns True if written."""
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


# Max rows below C to search for end-of-file (blank column M) when placing Q 0.
C_RANGE_ZERO_MAX_OFFSET = 100


def _resolve_q_zero_row_after_c(sheet, c_row: int) -> int:
    """
    Row for the Q 0 that ends the C range.

    Scan down from the row after C, up to C_RANGE_ZERO_MAX_OFFSET rows:
      - first row with no day flag in column M (end of trade list) -> place 0 there
      - if every row in the window still has M -> place 0 at C + 100
    """
    for offset in range(1, C_RANGE_ZERO_MAX_OFFSET + 1):
        row = c_row + offset
        if parse_m_dow(sheet.range(f"{COL_M}{row}").value) is None:
            return row
    return c_row + C_RANGE_ZERO_MAX_OFFSET


def _clear_stale_q_zeros_between(sheet, start_row_exclusive: int, end_row_exclusive: int) -> None:
    """Remove leftover Q stop 0s in (start, end) so Get_Closes does not stop early."""
    for r in range(start_row_exclusive + 1, end_row_exclusive):
        val = sheet.range(f"{COL_Q}{r}").value
        if _is_stop_flag(val) and not _is_q_chain_flag(val):
            sheet.range(f"{COL_Q}{r}").value = None
            print(f"  Cleared stale Q 0@{r} before new C-range terminator.")


def _place_q_zero_after_c(sheet, c_row: int) -> int:
    """
    Write the Q 0 that closes the C range (far terminator).
    Clears stale early 0s between C and the new terminator row.
    """
    zero_row = _resolve_q_zero_row_after_c(sheet, c_row)
    _clear_stale_q_zeros_between(sheet, c_row, zero_row)
    if _write_q_zero_safe(sheet, zero_row):
        print(f"  Q C-range terminator 0@{zero_row} (C@{c_row})")
    return zero_row


def _write_q_c_range(sheet, start_row: int, last_data_row: int) -> int:
    """Column Q: C at start; 0 at end-of-M / up to 100 rows below C."""
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
    - Place Q stop 0 via end-of-M / 100-row rule after C.
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
    max_row = _sheet_max_row(sheet)
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


def flag_prices_full_auto(
    *,
    earnings_file: str = LATEST_EARNINGS_FILE,
    today: Optional[date] = None,
) -> None:
    """Write P/Q flags for today's Full_Auto price jobs. No-op when nothing to flag."""
    today = today or date.today()
    if not is_trading_day(today):
        print(f"Flag_Prices_Full_Auto: {today} is not a trading day; skip.")
        return
    if not os.path.exists(earnings_file):
        print(f"Flag_Prices_Full_Auto: earnings file not found: {earnings_file}")
        return

    print(f"\nFlag_Prices_Full_Auto for {today}...")
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

        trade_dates = sorted({r.trade_date for r in all_rows})

        d_for_c = [d for d in trade_dates if d == today]
        d_for_m1 = [d for d in trade_dates if add_trading_days(d, 1) == today]
        d_for_m2 = [d for d in trade_dates if add_trading_days(d, 2) == today]

        if not d_for_c and not d_for_m1 and not d_for_m2:
            print(f"  No Full_Auto rows for C/M1/O/M2 jobs on {today}.")
            return

        wrote_m2_last: List[int] = []
        wrote_m1 = False
        wrote_c = False
        c_start_rows: List[int] = []

        # Write C first so M1/O can place P's 0 on the C row without inventing a Q 0.
        for d in d_for_c:
            rows = _rows_for_trade_date(all_rows, d)
            groups = _contiguous_groups(rows)
            # Snapshot empty-S starts before write (those become C rows).
            for g in groups:
                idx = _first_empty_dest_index(sheet, g, COL_S)
                if idx is not None:
                    c_start_rows.append(g[idx].row)
            c_rows = _flag_job_for_groups(
                sheet, groups, dest_col=COL_S, q_flag="C", job_label=f"C->S (D={d})"
            )
            if c_rows:
                wrote_c = True

        for d in d_for_m2:
            rows = _rows_for_trade_date(all_rows, d)
            groups = _contiguous_groups(rows)
            wrote_m2_last.extend(
                _flag_job_for_groups(
                    sheet, groups, dest_col=COL_V, q_flag="M2", job_label=f"M2->V (D={d})"
                )
            )

        for d in d_for_m1:
            rows = _rows_for_trade_date(all_rows, d)
            groups = _contiguous_groups(rows)
            m1_rows = _flag_job_for_groups(
                sheet,
                groups,
                dest_col=COL_T,
                p_flag="O",
                q_flag="M1",
                job_label=f"O->T + M1->U (D={d}, D+1={today})",
            )
            if m1_rows:
                wrote_m1 = True

        # Re-assert C flags in case anything still tried to place a 0 on those rows.
        for row in sorted(set(c_start_rows)):
            if not _is_q_c_flag(sheet.range(f"{COL_Q}{row}").value):
                sheet.range(f"{COL_Q}{row}").value = "C"
                print(f"  Restored Q C@{row} after M1/M2 pass.")

        # M2-only day: stop 0 only if it would not overwrite C/M1.
        if wrote_m2_last and not wrote_m1 and not wrote_c:
            zero_row = max(wrote_m2_last) + 1
            if _write_q_zero_safe(sheet, zero_row):
                print(f"  M2-only: wrote Q 0@{zero_row} to end M2 range.")

        # Strict layout QC — fail the run if invariants are broken.
        validate_flag_layout_strict(sheet)

        try:
            wb.save()
        except Exception:
            try:
                wb.save()
            except Exception as e:
                print(f"  Warning: could not save earnings workbook after flagging: {e}")
        print("Flag_Prices_Full_Auto: done.")
    finally:
        release_earnings_workbook(app, wb, owned_app=owned_app, owned_book=owned_book)