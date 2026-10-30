"""
Get Opens Auto — Automated opening prices into Latest Earnings by column K flags

Reads the Latest Earnings workbook (Trades sheet). Column K contains flags:
  - First "O" (letter): start; write opening prices to column T for each row.
  - Continue row by row until a "0" flag; then stop (do not process the 0 row).

For each processed row: ticker from column A, opening price from yfinance,
written to column T. Prints a summary of tickers written.

The Earnings file is opened and saved via xlwings (Excel) so that external
links and other Excel features are preserved. Close the workbook in Excel
before running.
"""

import os

import yfinance as yf

try:
    import xlwings as xw
except ImportError:
    xw = None

# ---------------------------------------------------------------------------
# Paths: Excel files live in repo root (Trading_Algo); script is in Old Scripts
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(os.path.dirname(_SCRIPT_DIR))

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SOURCE_FILE = os.path.join(_BASE_DIR, "! -- Latest Earnings Season.xlsx")
SOURCE_SHEET = "Trades"
HEADER_ROW = 3  # Data starts the row after headers (e.g. row 4)
COL_TICKER = "A"
COL_FLAG = "K"
COL_OPENING_PRICE = "T"
CHECK_END_ROW = 550  # Validate column K only through this row for a single O and single 0


def _is_start_flag(cell_value) -> bool:
    """True if cell is the letter O (start). Case-insensitive."""
    if cell_value is None:
        return False
    s = str(cell_value).strip().upper()
    return s == "O"


def _is_stop_flag(cell_value) -> bool:
    """True if cell is 0 (stop). Accepts numeric 0/0.0 or string '0'."""
    if cell_value is None:
        return False
    if isinstance(cell_value, (int, float)):
        return cell_value == 0
    return str(cell_value).strip() == "0"


def _normalize_ticker(cell_value):
    if cell_value is None:
        return None
    t = str(cell_value).strip().upper().replace(".", "-")
    return t if t else None


def _fetch_opens_batch(tickers: list[str]) -> dict[str, float | None]:
    """Fetch latest opening price for each ticker via yfinance (batch). Returns dict ticker -> price."""
    if not tickers:
        return {}
    unique = list(dict.fromkeys(tickers))
    try:
        data = yf.download(
            unique, period="5d", group_by="ticker", auto_adjust=True, progress=False, threads=True
        )
    except Exception:
        return {t: None for t in unique}
    if data.empty:
        return {t: None for t in unique}
    result = {}
    for t in unique:
        try:
            if len(unique) == 1 and "Open" in data.columns:
                open_series = data["Open"]
            elif hasattr(data.columns, "get_level_values") and t in data.columns.get_level_values(0):
                open_series = data[t]["Open"]
            else:
                result[t] = None
                continue
            val = open_series.iloc[-1]
            result[t] = round(float(val), 2) if val is not None else None
        except Exception:
            result[t] = None
    return result


def main():
    if xw is None:
        print("xlwings is not installed. Install it with: pip install xlwings")
        return
    if not os.path.exists(SOURCE_FILE):
        print(f"Source file not found: {SOURCE_FILE}")
        return

    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(os.path.abspath(SOURCE_FILE))
        try:
            sheet = wb.sheets[SOURCE_SHEET]
        except Exception:
            print(f"Sheet '{SOURCE_SHEET}' not found in {SOURCE_FILE}.")
            wb.close()
            return

        try:
            max_row = sheet.used_range.last_cell.row
        except Exception:
            max_row = 2000
        start_row = HEADER_ROW + 1

        # Ensure exactly one O and one 0 in column K through row 550
        count_o = 0
        count_zero = 0
        end_check = min(max_row, CHECK_END_ROW)
        for row in range(start_row, end_check + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_start_flag(cell_val):
                count_o += 1
            elif _is_stop_flag(cell_val):
                count_zero += 1
        if count_o != 1 or count_zero != 1:
            print("One clean range is not selected, please clean up your open range and try again.")
            wb.close()
            return

        # Find first "O" in column K
        first_o_row = None
        for row in range(start_row, max_row + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_start_flag(cell_val):
                first_o_row = row
                break

        if first_o_row is None:
            print("No 'O' flag found in column K. Nothing to process.")
            wb.close()
            return

        # From first O row until we hit 0: collect (row, ticker)
        to_process: list[tuple[int, str]] = []
        for row in range(first_o_row, max_row + 1):
            cell_val = sheet.range(f"{COL_FLAG}{row}").value
            if _is_stop_flag(cell_val):
                break
            ticker_raw = sheet.range(f"{COL_TICKER}{row}").value
            ticker = _normalize_ticker(ticker_raw)
            if not ticker:
                continue
            to_process.append((row, ticker))

        if not to_process:
            print("No tickers found in rows between 'O' and '0' in column K.")
            wb.close()
            return

        tickers_to_fetch = [t for _, t in to_process]
        print(f"Fetching opening prices for {len(tickers_to_fetch)} ticker(s)...")
        prices = _fetch_opens_batch(tickers_to_fetch)

        written: list[str] = []
        for row, ticker in to_process:
            price = prices.get(ticker)
            if price is not None:
                sheet.range(f"{COL_OPENING_PRICE}{row}").value = price
                written.append(ticker)
            else:
                print(f"  Warning: no opening price for {ticker} (row {row})")

        wb.save()
        wb.close()
        print("Opening prices written to Latest Earnings (saved via Excel).")
        print(f"\nTickers written to column {COL_OPENING_PRICE}: {', '.join(written)}")
    finally:
        if app is not None:
            app.quit()


if __name__ == "__main__":
    main()
