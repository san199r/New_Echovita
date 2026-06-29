#!/usr/bin/env python3
"""
Excel Cleanup Tool: Normalize 'Dod' / 'DOD' dates to mm/dd/yyyy.

What it does
------------
- Opens a file browser so you can select an Excel workbook (.xlsx, .xls, .xlsm).
- Looks for a column named 'Dod' / 'DOD' (case-insensitive, ignoring surrounding spaces/dots).
- Parses date strings assuming **dd-mm-yyyy** (e.g., "27-10-2025" = 27 Oct 2025).
- Outputs the column strictly formatted as 'mm/dd/yyyy' (e.g., 10/27/2025).
- Saves a new file next to the original with suffix "_cleaned.xlsx".

Usage
-----
- From PowerShell / cmd:
    D:\Workspace\venv\Scripts\python.exe D:\Workspace\Echovita_Dates_Cleanup_Tool.py
- If pandas/openpyxl are missing, install with:
    pip install pandas openpyxl
"""

import os
import re
import sys
from datetime import datetime
from typing import Optional

print("Excel Cleanup Tool - build 2025-12-05")  # simple sanity check of which script is running

try:
    import pandas as pd
except ImportError:
    print("This tool requires 'pandas'. Install it with: pip install pandas openpyxl")
    sys.exit(1)

# Removed tkinter dependencies

DATE_PATTERN = re.compile(r'^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\s*$')


def normalize_year(y: int) -> int:
    """Normalize 2-digit years to 2000-2099; pass through 4-digit years."""
    if 0 <= y <= 99:
        return 2000 + y
    return y


def parse_mixed_date(value) -> Optional[datetime]:
    """
    Parse a date assuming dd-mm-yyyy (or dd/mm/yyyy) and return a datetime.

    Rules:
    - Strings matching d{1,2}[-/]d{1,2}[-/]d{2,4} are interpreted strictly as dd-mm-yyyy.
    - Existing datetime-like values are passed through.
    - Excel serial numbers are parsed via Excel's 1899-12-30 origin.
    - For other strings, we try pandas with dayfirst=True only.

    Returns None if parsing fails.
    """
    if pd.isna(value):
        return None

    # Already datetime-like?
    if isinstance(value, (datetime, pd.Timestamp)):
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.to_pydatetime()

    # Excel serial or numeric-like?
    try:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notna(dt):
                return dt.to_pydatetime()
    except Exception:
        pass

    s = str(value).strip()
    m = DATE_PATTERN.match(s)
    if m:
        d, mth, y = m.groups()
        d, mth, y = int(d), int(mth), int(y)
        y = normalize_year(y)
        try:
            # dd-mm-yyyy assumption
            return datetime(y, mth, d)
        except ValueError:
            return None

    # Fallback: interpret with dayfirst=True (dd-mm)
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    except Exception:
        pass

    return None


def format_mm_dd_yyyy(dt: datetime) -> str:
    """Return date as mm/dd/yyyy."""
    return dt.strftime("%m/%d/%Y")


def find_dod_column(df: pd.DataFrame) -> Optional[str]:
    """Find the 'Dod' / 'DOD' column (case-insensitive, ignoring spaces and dots)."""
    normalized = [str(c).strip().lower() for c in df.columns]
    for original, norm in zip(df.columns, normalized):
        if norm == "dod":
            return original
    # Also try softer matches like 'do d', 'd.o.d', etc.
    for original, norm in zip(df.columns, normalized):
        if norm.replace(".", "").replace(" ", "") == "dod":
            return original
    return None


def clean_workbook(path: str) -> str:
    """Clean all sheets in the workbook and save to *_cleaned.xlsx."""
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        raise RuntimeError(f"Failed to open Excel file:\n{e}")

    if not xl.sheet_names:
        raise RuntimeError(
            "The workbook has no visible worksheets. "
            "Open it in Excel and make sure it has at least one normal (visible) sheet."
        )

    print("Input sheets:", xl.sheet_names)

    out_path = os.path.splitext(path)[0] + "_Cleaned.xlsx"

    # Use openpyxl engine but force all sheets visible before saving
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name in xl.sheet_names:
            # Read as strings to avoid auto-coercion surprises
            df = xl.parse(sheet_name=sheet_name, dtype=str)

            dod_col = find_dod_column(df)
            if dod_col is not None:
                # Parse and format safely: handle None/NaT explicitly
                parsed = df[dod_col].apply(parse_mixed_date)
                formatted = parsed.apply(
                    lambda d: format_mm_dd_yyyy(d)
                    if (d is not None and not pd.isna(d))
                    else ""
                )
                df[dod_col] = formatted

            # Write the sheet (modified or unchanged)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # ---- CRITICAL FIX: force all sheets to be visible before save ----
        wb = writer.book
        sheets_states = [(ws.title, getattr(ws, "sheet_state", "unknown")) for ws in wb.worksheets]
        print("Sheet states before forcing visible:", sheets_states)

        for ws in wb.worksheets:
            ws.sheet_state = "visible"

        sheets_states_after = [(ws.title, ws.sheet_state) for ws in wb.worksheets]
        print("Sheet states after forcing visible:", sheets_states_after)

    return out_path


def main():
    # If passed via command line
    if len(sys.argv) > 1:
        in_path = sys.argv[1]
    else:
        in_path = os.path.join(os.getcwd(), "workspace", "output", "Echovita_Family_Links_Tool_Output.xlsx")

    if not os.path.exists(in_path):
        print(f"Input file not found: {in_path}")
        sys.exit(1)

    try:
        print("Processing file:", in_path)
        out_path = clean_workbook(in_path)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

    print(f"Cleaned file saved:\n{out_path}")


if __name__ == "__main__":
    main()
