import re
import logging
from datetime import datetime, date
from typing import Tuple, List
from dateutil.relativedelta import relativedelta


def coordinate_to_tuple(cell_ref: str) -> Tuple[int, int]:
    """Converts an Excel cell reference (e.g., 'A1', 'B10') to a (row, column) tuple."""
    match = re.match(r"([A-Za-z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_str, row_str = match.groups()
    col_num = 0
    for i, char in enumerate(reversed(col_str)):
        col_num += (ord(char.upper()) - 64) * (26 ** i)
    return int(row_str), col_num

def get_fiscal_quarter_and_month(date_obj: date) -> Tuple[int, str, str, str]:
    """
    Determines the fiscal year, fiscal quarter (Q1, Q2, ...), overall fiscal month (M1, M2, ... M12),
    and month within the quarter (M1, M2, M3) for a given date, with the fiscal year starting in August.
    """
    fiscal_start_month = 8  # August

    if date_obj.month >= fiscal_start_month:
        fiscal_year = date_obj.year + 1
    else:
        fiscal_year = date_obj.year

    fiscal_month_overall_num = ((date_obj.month - fiscal_start_month) % 12) + 1
    fiscal_quarter_num = (fiscal_month_overall_num - 1) // 3 + 1
    fiscal_month_in_quarter_num = (fiscal_month_overall_num - 1) % 3 + 1

    return fiscal_year, f"Q{fiscal_quarter_num}", f"M{fiscal_month_overall_num}", f"M{fiscal_month_in_quarter_num}"

def find_dynamic_sheets(sheet_names: List[str], fiscal_quarter_str: str, fiscal_month_overall_str: str, fiscal_month_in_quarter_str: str) -> List[str]:
    """
    Dynamically finds sheet names based on the fiscal quarter, overall fiscal month,
    month within quarter, and required patterns.
    """
    required_patterns = [
        r"^Margins Scenarios$",
        rf"^{fiscal_month_in_quarter_str} .*Exec View$",
        rf"^{fiscal_quarter_str} {fiscal_month_in_quarter_str} .*Comparisons$",
        rf"^{fiscal_quarter_str} Commit$",
    ]

    matched_sheets = []
    for pattern in required_patterns:
        matched = False
        for sheet_name in sheet_names:
            if re.match(pattern, sheet_name, re.IGNORECASE):
                matched_sheets.append(sheet_name)
                matched = True
                break
        if not matched:
            logging.warning(f"No match found for pattern: {pattern}")

    return matched_sheets

def get_dynamic_filename_components() -> Tuple[str, str]:
    """
    Calculates the expected YYYYMM prefix for the source Excel file
    and today's date in YYYYMMDD format for dynamic file searching.
    """
    now = datetime.now()

    # Logic to determine future month offset for the source Excel filename prefix
    if now.month == 10 and now.year == 2025:
        future_month_offset = 4
    else:
        future_month_offset = 5

    target_future_date = now + relativedelta(months=future_month_offset)
    future_prefix_yyyymm = f"{target_future_date.year}{target_future_date.month:02d}"

    today_yyyymmdd = now.strftime("%Y%m%d")

    return future_prefix_yyyymm, today_yyyymmdd

def iterate_all_shapes(shapes):
    """Recursively yields all shapes within a slide or group shape."""
    for shp in shapes:
        yield shp
        if hasattr(shp, "shapes"):
            yield from iterate_all_shapes(shp.shapes)