"""
Computes per-instructor fee summaries from parsed transaction data.

Calculations per instructor (grouped by income_center_group → item):
    Sum Amt After Disc = sum of amt_after_disc
    Comm              = Sum Amt After Disc × commission_rate  (default 5 %)
    Total             = Sum Amt After Disc − Comm
    EWT               = Total × ISP tax_rate
    Final Total       = Total − EWT

Returns a dict keyed by group name, each value being a result dict:
    {
        "rows": DataFrame,       # one row per instructor (sorted by name)
        "grand_total": dict,     # aggregated totals row
        "date_min": datetime,
        "date_max": datetime,
    }
"""
from datetime import datetime
import re as _re

import pandas as pd


COMMISSION_RATE_DEFAULT = 0.05


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def summarise(
    df: pd.DataFrame,
    isp_dict: dict,
    commission_rate: float = COMMISSION_RATE_DEFAULT,
    aliases: dict | None = None,
) -> tuple[dict, list[str]]:
    """
    Args:
        df:               flat DataFrame from raw_data_reader.read()
        isp_dict:         dict from isp_reader.load()
        commission_rate:  fixed commission fraction (e.g. 0.05 for 5 %)
        aliases:          {pos_name_lower: isp_display_name} saved mappings;
                          None means no aliases.

    Returns:
        (summaries, warnings)
        summaries: {group_name: {"rows": DataFrame, "grand_total": dict,
                                 "date_min": datetime, "date_max": datetime}}
    """
    if aliases is None:
        aliases = {}
    warnings: list = []
    summaries: dict = {}

    for (group_name, center_name), group_df in df.groupby(
        ["income_center_group", "income_center"], sort=True
    ):
        rows = []

        for instructor, inst_df in group_df.groupby("item", sort=True):
            instructor_stripped = str(instructor).strip()
            if not instructor_stripped:
                continue

            # ISP lookup — apply alias first, then fall back to exact match
            lookup_key = instructor_stripped.lower()
            if lookup_key in aliases and aliases[lookup_key] is not None:
                lookup_key = aliases[lookup_key].lower()
            isp_info = isp_dict.get(lookup_key)
            if isp_info is None:
                warnings.append(
                    f"'{instructor_stripped}' was not found in the ISP list. "
                    "EWT set to 0."
                )
                ewt_rate = 0.0
            else:
                ewt_rate = isp_info["tax_rate"]

            min_unit_price = _r2(float(inst_df["unit_price"].min()))
            sum_qty        = float(inst_df["qty"].sum())   # keep as float to preserve 0.5 etc.
            sum_amt        = _r2(float(inst_df["amt_after_disc"].sum()))
            comm           = _r2(sum_amt * commission_rate)
            total          = _r2(sum_amt - comm)
            ewt            = _r2(total * ewt_rate)
            final_total    = _r2(total - ewt)

            rows.append({
                "item":             instructor_stripped,
                "min_unit_price":   min_unit_price,
                "sum_qty":          sum_qty,
                "sum_amt":          sum_amt,
                "comm":             comm,
                "total":            total,
                "ewt_rate":         ewt_rate,
                "ewt":              ewt,
                "final_total":      final_total,
            })

        if not rows:
            continue

        summary_df = pd.DataFrame(rows)

        grand_total = {
            "item":           "Grand Total",
            "min_unit_price": None,
            "sum_qty":        float(summary_df["sum_qty"].sum()),   # keep float
            "sum_amt":        _r2(float(summary_df["sum_amt"].sum())),
            "comm":           _r2(float(summary_df["comm"].sum())),
            "total":          _r2(float(summary_df["total"].sum())),
            "ewt_rate":       None,
            "ewt":            _r2(float(summary_df["ewt"].sum())),
            "final_total":    _r2(float(summary_df["final_total"].sum())),
        }

        # Date range from the raw transactions in this group
        dates = pd.to_datetime(group_df["date"], errors="coerce").dropna()
        date_min = dates.min().to_pydatetime() if not dates.empty else datetime.today()
        date_max = dates.max().to_pydatetime() if not dates.empty else datetime.today()

        summaries[center_name] = {
            "rows":               summary_df,
            "grand_total":        grand_total,
            "date_min":           date_min,
            "date_max":           date_max,
            "income_center_group": group_name,
            "income_center":       center_name,
        }

    return summaries, warnings


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _r2(value: float) -> float:
    """Round to 2 decimal places using standard rounding."""
    return round(value, 2)


def format_date_range(date_min: datetime, date_max: datetime) -> str:
    """
    Format two dates as a compact range string.
        Same month/year:  "May 1-8, 2026"
        Different months: "Dec 28, 2025 - Jan 3, 2026"
    """
    if date_min.year == date_max.year and date_min.month == date_max.month:
        return f"{date_min.strftime('%b')} {date_min.day}-{date_max.day}, {date_max.year}"
    # Use .day (int) to avoid platform-specific strftime flags like %-d
    return (
        f"{date_min.strftime('%b')} {date_min.day}, {date_min.year} - "
        f"{date_max.strftime('%b')} {date_max.day}, {date_max.year}"
    )


def group_to_title(center_name: str) -> str:
    """
    Convert an income_center name to a report title.
        "Archery - Instructor Fees"  → "Archery Instructor Fee"
        "Pickleball Coach"           → "Pickleball Coach Instructor Fee"
        "ARCHERY"                    → "Archery Instructor Fee"  (legacy)
    """
    cleaned = _re.sub(
        r'\s*-?\s*instructor fees?\s*$', '', center_name, flags=_re.IGNORECASE
    ).strip(' -')
    label = cleaned.title() if cleaned else center_name.title()
    return label + " Instructor Fee"


def tab_label(center_name: str) -> str:
    """
    Short label for a notebook tab (strips 'Instructor Fees' suffix).
        "Archery - Instructor Fees"  → "Archery"
        "Pickleball Coach"           → "Pickleball Coach"
    """
    cleaned = _re.sub(
        r'\s*-?\s*instructor fees?\s*$', '', center_name, flags=_re.IGNORECASE
    ).strip(' -')
    return cleaned.title() if cleaned else center_name.title()
