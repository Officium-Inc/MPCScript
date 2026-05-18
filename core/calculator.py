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

import pandas as pd


COMMISSION_RATE_DEFAULT = 0.05


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def summarise(
    df: pd.DataFrame,
    isp_dict: dict,
    commission_rate: float = COMMISSION_RATE_DEFAULT,
) -> tuple[dict, list[str]]:
    """
    Args:
        df:               flat DataFrame from raw_data_reader.read()
        isp_dict:         dict from isp_reader.load()
        commission_rate:  fixed commission fraction (e.g. 0.05 for 5 %)

    Returns:
        (summaries, warnings)
        summaries: {group_name: {"rows": DataFrame, "grand_total": dict,
                                 "date_min": datetime, "date_max": datetime}}
    """
    warnings: list = []
    summaries: dict = {}

    for group_name, group_df in df.groupby("income_center_group", sort=True):
        rows = []

        for instructor, inst_df in group_df.groupby("item", sort=True):
            instructor_stripped = str(instructor).strip()
            if not instructor_stripped:
                continue

            # ISP lookup — case-insensitive
            isp_info = isp_dict.get(instructor_stripped.lower())
            if isp_info is None:
                warnings.append(
                    f"'{instructor_stripped}' was not found in the ISP list. "
                    "EWT set to 0."
                )
                ewt_rate = 0.0
            else:
                ewt_rate = isp_info["tax_rate"]

            min_unit_price = _r2(float(inst_df["unit_price"].min()))
            sum_qty        = int(inst_df["qty"].sum())
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
            "sum_qty":        int(summary_df["sum_qty"].sum()),
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

        summaries[group_name] = {
            "rows":        summary_df,
            "grand_total": grand_total,
            "date_min":    date_min,
            "date_max":    date_max,
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


def group_to_title(group_name: str) -> str:
    """
    Convert a group name to a report title.
        "ARCHERY"           → "Archery Instructor Fee"
        "GOLF DRIVING RANGE"→ "Golf Driving Range Instructor Fee"
    """
    return group_name.title() + " Instructor Fee"
