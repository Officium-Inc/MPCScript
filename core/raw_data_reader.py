"""
Parses the raw POS Sales-Income Center Excel export.

The POS file has a formatted report structure (not a simple flat table):
  - Report header rows  (company name, title, date range, etc.)
  - "Income Center Group : ARCHERY"  section header
  - "Income Center : Archery - Instructor Fees"  section header
  - Column header row  (Date | Receipt No. | Pax | Item | ...)
  - Data rows          (one per transaction; Date column holds a datetime)
  - "Income Center Total : ..."  subtotal row  → skipped
  - "Income Center Group Total : ..."  subtotal row  → skipped
  - Above pattern repeats for each group / center

Returns a flat DataFrame with all transaction columns plus two added columns:
    income_center_group   e.g. "ARCHERY"
    income_center         e.g. "Archery - Instructor Fees"
"""
import re
from datetime import datetime
from pathlib import Path
from typing import Union

import pandas as pd

# Matches "5/4/2026", "05/04/2026", "12/31/2025"
_DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
# Matches "2026-05-04"
_DATE_ISO_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read(filepath: str) -> tuple[pd.DataFrame, list[str]]:
    """
    Parse the POS raw data file.

    Returns:
        (DataFrame, warnings_list)

    Raises:
        ValueError: if no transaction rows are found.
    """
    path = Path(filepath)
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"

    # Read without imposing headers; let pandas infer types so dates come back
    # as Timestamp objects and numbers as floats.
    raw: pd.DataFrame = pd.read_excel(filepath, engine=engine, header=None)

    col_map: dict = {}          # logical_name → column index
    current_group: str = ""
    current_center: str = ""
    records: list = []
    warnings: list = []

    for row_idx in range(len(raw)):
        row = raw.iloc[row_idx]

        # Build a list of clean string representations for text-based detection
        str_cells = _str_cells(row)
        joined = " ".join(str_cells).lower()

        # --- Skip completely empty rows ---
        if not any(s.strip() for s in str_cells):
            continue

        # --- Skip all "* Total *" summary rows ---
        if "total" in joined and ("income center" in joined or "grand" in joined):
            continue

        # --- Income Center Group header (check BEFORE plain "income center") ---
        if "income center group" in joined and ":" in joined:
            name = _extract_after_colon(str_cells, "income center group")
            if name:
                current_group = name
            continue

        # --- Income Center header ---
        if "income center" in joined and "group" not in joined and ":" in joined:
            name = _extract_after_colon(str_cells, "income center")
            if name:
                current_center = name
            continue

        # --- Column header row ---
        if _is_header_row(str_cells):
            col_map = _build_col_map(str_cells)
            continue

        # --- Data row ---
        if not col_map:
            continue

        date_idx = col_map.get("date")
        if date_idx is None:
            continue

        date_val = row.iloc[date_idx]
        if not _is_date(date_val):
            continue

        rec = _extract_record(row, str_cells, col_map, current_group, current_center)
        if rec:
            records.append(rec)

    if not records:
        raise ValueError(
            "No transaction rows were found in the file.\n"
            "Please verify that the file is the correct POS Sales-Income Center export."
        )

    df = pd.DataFrame(records)

    # Warn about rows where instructor name is blank
    blank_items = df[df["item"].str.strip() == ""]
    if not blank_items.empty:
        warnings.append(
            f"{len(blank_items)} row(s) had a blank Item (instructor) name and were included."
        )

    return df, warnings


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _str_cells(row: pd.Series) -> list:
    """Convert a DataFrame row to a list of stripped strings (NaN → '')."""
    result = []
    for val in row:
        if val is None:
            result.append("")
            continue
        try:
            if pd.isna(val):
                result.append("")
                continue
        except (TypeError, ValueError):
            pass
        result.append(str(val).strip())
    return result


def _is_date(val) -> bool:
    """Return True if val is a date (Timestamp, datetime, or date-string) in a plausible year range."""
    if val is None:
        return False
    try:
        if pd.isna(val):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(val, (pd.Timestamp, datetime)):
        try:
            return 2015 <= val.year <= 2035
        except Exception:
            return False
    if isinstance(val, str):
        s = val.strip()
        if _DATE_RE.match(s):
            try:
                return 2015 <= datetime.strptime(s, "%m/%d/%Y").year <= 2035
            except ValueError:
                pass
        if _DATE_ISO_RE.match(s):
            try:
                return 2015 <= datetime.strptime(s, "%Y-%m-%d").year <= 2035
            except ValueError:
                pass
    return False


def _is_header_row(cells: list) -> bool:
    """A column-header row contains 'date', 'item', and 'qty' as cell values."""
    lower_set = {c.lower().strip() for c in cells if c}
    return "date" in lower_set and "item" in lower_set and "qty" in lower_set


def _build_col_map(cells: list) -> dict:
    """Map logical column names to their integer index positions."""
    col_map: dict = {}
    for i, cell in enumerate(cells):
        c = cell.lower().strip()
        if c == "date":
            col_map["date"] = i
        elif "receipt" in c:
            col_map["receipt_no"] = i
        elif c == "pax":
            col_map["pax"] = i
        elif c == "item":
            col_map["item"] = i
        elif c == "guest" and "pay" not in c:
            col_map["guest"] = i
        elif "pay" in c and "guest" in c:
            col_map["pay_guest"] = i
        elif "unit" in c and "price" in c:
            col_map["unit_price"] = i
        elif c == "qty":
            col_map["qty"] = i
        elif "gross" in c:
            col_map["gross_price"] = i
        elif "amt" in c and "disc" in c:
            # "Amt. After Disc." — check before plain "disc"
            col_map["amt_after_disc"] = i
        elif "disc" in c and "amt" not in c:
            col_map["disc"] = i
        elif c == "svc":
            col_map["svc"] = i
        elif c == "tax":
            col_map["tax"] = i
        elif c == "total":
            col_map["total"] = i
    return col_map


def _extract_after_colon(cells: list, keyword: str) -> str:
    """
    Join all non-empty cells, find keyword, then return the text after ':'.
    Handles both 'Income Center Group : ARCHERY' in one cell and
    spread across multiple cells.
    """
    full = " ".join(c for c in cells if c.strip())
    lower = full.lower()
    idx = lower.find(keyword.lower())
    if idx == -1:
        return ""
    colon_idx = full.find(":", idx)
    if colon_idx == -1:
        return ""
    return full[colon_idx + 1:].strip()


def _get_float(row: pd.Series, idx: Union[int, None]) -> float:
    """Safely extract a float from a row by column index."""
    if idx is None or idx >= len(row):
        return 0.0
    val = row.iloc[idx]
    if val is None:
        return 0.0
    try:
        if pd.isna(val):
            return 0.0
    except (TypeError, ValueError):
        pass
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _get_str(str_cells: list, idx: Union[int, None]) -> str:
    if idx is None or idx >= len(str_cells):
        return ""
    return str_cells[idx]


def _extract_record(
    row: pd.Series,
    str_cells: list,
    col_map: dict,
    group: str,
    center: str,
) -> dict:
    date_val = row.iloc[col_map["date"]]
    # Normalise to Python datetime for consistent handling
    if isinstance(date_val, pd.Timestamp):
        date_val = date_val.to_pydatetime()
    elif isinstance(date_val, str):
        s = date_val.strip()
        try:
            date_val = datetime.strptime(s, "%m/%d/%Y")
        except ValueError:
            try:
                date_val = datetime.strptime(s, "%Y-%m-%d")
            except ValueError:
                date_val = datetime.today()

    # Receipt No. is stored as text (e.g. '047203') — preserve the original string
    receipt_raw = row.iloc[col_map["receipt_no"]] if "receipt_no" in col_map else ""
    receipt_str = str(receipt_raw).strip() if receipt_raw not in ("", None) else ""
    # Drop trailing ".0" if pandas loaded it as a float string
    if receipt_str.endswith(".0"):
        receipt_str = receipt_str[:-2]

    return {
        "date": date_val,
        "receipt_no": receipt_str,
        "pax": _get_str(str_cells, col_map.get("pax")),
        "item": _get_str(str_cells, col_map.get("item")),
        "guest": _get_str(str_cells, col_map.get("guest")),
        "unit_price": _get_float(row, col_map.get("unit_price")),
        "qty": _get_float(row, col_map.get("qty")),
        "gross_price": _get_float(row, col_map.get("gross_price")),
        "disc": _get_float(row, col_map.get("disc")),
        "amt_after_disc": _get_float(row, col_map.get("amt_after_disc")),
        "svc": _get_float(row, col_map.get("svc")),
        "tax": _get_float(row, col_map.get("tax")),
        "total": _get_float(row, col_map.get("total")),
        "income_center_group": group,
        "income_center": center,
    }
