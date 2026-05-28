"""
Reads the ISP reference list Excel file.
Returns a dict keyed by lowercase instructor name for fast lookup.
"""
from pathlib import Path

import pandas as pd


def load(filepath: str) -> dict:
    """
    Load ISP list from .xlsx or .xls file.

    Returns:
        {
            "joan tabanag": {
                "display_name": "Joan Tabanag",
                "tax_rate": 0.05,
                "tin_no": "227-123-513-000",
                "sworn_decl": "B",
                "vat_registered": "No",
                "acct_no": "ISP0181",
            },
            ...
        }
    Raises:
        ValueError: if the file is empty or cannot be parsed.
    """
    path = Path(filepath)
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"

    # --- Step 1: find the actual column-header row ---
    # The file may have several title rows before the real headers
    # (e.g. "S&A ISP INVOICING", date range, totals row …).
    raw = pd.read_excel(filepath, engine=engine, header=None, nrows=20)
    header_row_idx = _find_header_row(raw)
    if header_row_idx == -1:
        raise ValueError(
            "This file does not appear to be an ISP Reference List.\n"
            "Expected a row containing headers such as "
            "'ISP Name', 'Tax Rate', 'TIN No.', or 'Sworn Declaration'."
        )

    # --- Step 2: re-read using the detected header row ---
    df = pd.read_excel(filepath, engine=engine, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = _map_columns(df.columns.tolist())

    result: dict = {}
    for _, row in df.iterrows():
        raw_name = str(row.get(col_map.get("isp_name", ""), "")).strip()
        if not raw_name or raw_name.lower() == "nan":
            continue

        tax_rate = _parse_rate(str(row.get(col_map.get("tax_rate", ""), "0")))

        def _get(key: str) -> str:
            val = row.get(col_map.get(key, ""), "")
            return "" if pd.isna(val) else str(val).strip()

        result[raw_name.lower()] = {
            "display_name": raw_name,
            "tax_rate": tax_rate,
            "tin_no": _get("tin_no"),
            "sworn_decl": _get("sworn_decl"),
            "vat_registered": _get("vat_registered"),
            "acct_no": _get("acct_no"),
        }

    if not result:
        raise ValueError(
            "ISP list appears empty or could not be parsed. "
            "Ensure it has columns: ISP NAME, Tax Rate."
        )

    return result


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_HEADER_KEYWORDS = {"isp name", "tax rate", "sworn", "tin no", "acct no"}


def _find_header_row(raw_df: pd.DataFrame) -> int:
    """Return the 0-based row index of the ISP column-header row, or -1 if not found."""
    for i, row in raw_df.iterrows():
        cells = {str(v).lower().strip() for v in row if str(v).lower() != "nan"}
        matches = sum(1 for kw in _HEADER_KEYWORDS if any(kw in c for c in cells))
        if matches >= 2:
            return int(i)
    return -1


# Keep for backward compatibility (used nowhere else, but harmless)
def _validate_isp_headers(columns: list) -> None:
    pass


def _map_columns(columns: list) -> dict:
    """Map logical names to actual column names using keyword matching."""
    col_map: dict = {}
    for col in columns:
        cl = col.lower()
        if "isp name" in cl or cl == "isp name":
            col_map["isp_name"] = col
        elif "tax rate" in cl:
            col_map["tax_rate"] = col
        elif "tin" in cl:
            col_map["tin_no"] = col
        elif "sworn" in cl:
            col_map["sworn_decl"] = col
        elif "vat" in cl:
            col_map["vat_registered"] = col
        elif "acct" in cl:
            col_map["acct_no"] = col

    # Positional fallbacks when column names are non-standard
    if "isp_name" not in col_map and len(columns) >= 3:
        col_map["isp_name"] = columns[2]   # Column C is typically ISP NAME
    if "tax_rate" not in col_map and len(columns) >= 6:
        col_map["tax_rate"] = columns[5]   # Column F is typically Tax Rate

    return col_map


def _parse_rate(value: str) -> float:
    """Parse '5%' → 0.05, '0.05' → 0.05, '5' → 0.05, blank/nan → 0.0."""
    value = str(value).replace("%", "").strip()
    if not value or value.lower() == "nan":
        return 0.0
    try:
        rate = float(value)
        if rate != rate:   # IEEE NaN guard
            return 0.0
        return rate / 100.0 if rate > 1.0 else rate
    except ValueError:
        return 0.0
