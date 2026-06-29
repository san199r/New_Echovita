#!/usr/bin/env python3
"""
Excel combine tool

- Input: pick one or MORE Excel files (.xlsx/.xlsm) OR pass them as CLI args.
- Output:
    • ALWAYS saved as ONE file next to the FIRST input file: 'Combined_Excel.xlsx'

Rules implemented:
  • Ignore 'Sheet1'
  • Use header A..H and O..AO from sheet '1' (fallback to lowest numeric sheet if '1' is missing)
    - Header is taken from the FIRST input file only (so all files align to one header)
  • Append data A..H and O..AO from sheets '1','2','3',... (row 2 onward; row 1 is header)
  • Convert date/datetime values to text 'dd-mm-YYYY' (no '00:00:00')
  • Renumber S.No. in the first column as 1..N (after combining ALL files)
  • Remove hyperlink behavior/blue styling in columns D and H
  • Print per-sheet status lines like:
      - Added Sheet named "1" data into output (rows: X)
      - No data in Sheet Named "2"

Requires:
  pip install pandas openpyxl xlsxwriter tqdm
"""

import os
import sys
import re
import pandas as pd
from tqdm import tqdm


# Removed popup and file picker functions


# ---------- Helpers ----------
def collect_numeric_sheets(xls):
    """Return numeric sheet names (e.g., '1','2','3',...) excluding 'Sheet1', sorted as integers."""
    numeric = []

    for s in xls.sheet_names:
        if s == "Sheet1":
            continue

        if re.fullmatch(r"\d+", str(s).strip()):
            numeric.append(str(s).strip())

    numeric.sort(key=lambda x: int(x))
    return numeric


def normalize_dates_to_text(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert any datetime-like columns to 'dd-mm-YYYY' text.
    Also tries to coerce likely date columns by name.
    Explicitly handles dd-mm-YYYY inputs to avoid pandas dayfirst warnings.
    """
    df = df.copy()

    # 1) Pure datetime columns -> format to dd-mm-YYYY text
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%d-%m-%Y")

    # 2) Likely date columns by name -> attempt coercion & format
    likely_date_cols = {"dod", "date of death", "date", "dob", "doj", "doa"}

    for col in df.columns:
        if isinstance(col, str) and col.strip().lower() in likely_date_cols:
            series = df[col].astype(str)

            # Detect exact dd-mm-YYYY
            mask_ddmmyyyy = series.str.match(r"^\s*\d{1,2}-\d{1,2}-\d{4}\s*$", na=False)

            # Parse exact dd-mm-YYYY safely
            parsed_exact = pd.to_datetime(
                series.where(mask_ddmmyyyy, None),
                errors="coerce",
                format="%d-%m-%Y"
            )

            # Parse everything else with dayfirst=True
            parsed_fallback = pd.to_datetime(
                series.where(~mask_ddmmyyyy, None),
                errors="coerce",
                dayfirst=True
            )

            parsed = parsed_exact.fillna(parsed_fallback)

            # Keep original text where parsing fails
            as_text = parsed.dt.strftime("%d-%m-%Y")
            df[col] = as_text.where(~parsed.isna(), series)

    return df


def make_writer(out_path: str):
    """
    Create a pandas ExcelWriter that disables URL auto-detection
    so text that looks like a URL won't turn blue/underlined.
    Supports both newer and older pandas.
    """
    try:
        # pandas >= 2.0
        return pd.ExcelWriter(
            out_path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}}
        )
    except TypeError:
        # pandas 1.x fallback
        return pd.ExcelWriter(
            out_path,
            engine="xlsxwriter",
            options={"strings_to_urls": False}
        )


def resolve_single_output_path(first_input_path: str) -> str:
    """ALWAYS write one combined file to workspace/output."""
    output_dir = os.path.join(os.getcwd(), "workspace", "output")
    os.makedirs(output_dir, exist_ok=True)
    return os.path.join(output_dir, "Combined_Excel.xlsx")


def read_header_columns_from_first_file(first_path: str):
    """
    Use header A..H and O..AO from sheet '1' in the FIRST file.
    Fallback to lowest numeric sheet if '1' is missing.
    Returns (header_sheet_name, columns_list).
    """
    xls = pd.ExcelFile(first_path)
    numeric_sheets = collect_numeric_sheets(xls)

    if not numeric_sheets:
        raise ValueError("FIRST file has no numeric sheets (e.g., '1','2','3', ...) after ignoring 'Sheet1'.")

    header_sheet = "1" if "1" in numeric_sheets else numeric_sheets[0]

    header_df = pd.read_excel(
        first_path,
        sheet_name=header_sheet,
        header=0,
        usecols="A:H,O:AO"
    )

    return header_sheet, list(header_df.columns)


def combine_all_files(input_paths, out_path):
    """
    Combine numeric sheets from ALL input files into ONE output file.
    Header columns are taken from FIRST input file only.
    """
    first_header_sheet, header_cols = read_header_columns_from_first_file(input_paths[0])

    combined = pd.DataFrame(columns=header_cols)
    status_msgs = []

    for file_idx, src_path in enumerate(tqdm(input_paths, desc="Processing files", unit="file"), start=1):

        status_msgs.append("-" * 80)
        status_msgs.append(f'File {file_idx}/{len(input_paths)}: {src_path}')

        try:
            xls = pd.ExcelFile(src_path)
        except Exception as e:
            status_msgs.append(f'❌ Failed to open file: {e}')
            continue

        numeric_sheets = collect_numeric_sheets(xls)

        if not numeric_sheets:
            status_msgs.append('❌ No numeric sheets found in this file (ignored).')
            continue

        for s in tqdm(numeric_sheets, desc=f"Sheets in file {file_idx}", unit="sheet", leave=False):
            try:
                df = pd.read_excel(
                    src_path,
                    sheet_name=s,
                    header=0,
                    usecols="A:H,O:AO"
                )
            except Exception as e:
                status_msgs.append(f'❌ Failed to read Sheet Named "{s}": {e}')
                continue

            # Force same columns as first file header
            df = df.reindex(columns=header_cols)

            # Drop fully empty rows
            df = df.dropna(how="all")

            if df.empty:
                status_msgs.append(f'No data in Sheet Named "{s}"')
                continue

            df = normalize_dates_to_text(df)
            combined = pd.concat([combined, df], ignore_index=True)

            status_msgs.append(f'Added Sheet named "{s}" data into output (rows: {len(df)})')

    # Final cleanup
    combined = combined.dropna(how="all")

    # Renumber S.No. in first column (after ALL files combined)
    if not combined.empty:
        first_col_name = combined.columns[0]
        combined[first_col_name] = range(1, len(combined) + 1)

    # Write one output file
    writer = make_writer(out_path)
    combined.to_excel(writer, sheet_name="Combined", index=False)

    # Remove hyperlink styling/behavior specifically on columns D and H
    workbook = writer.book
    worksheet = writer.sheets["Combined"]

    plain_fmt = workbook.add_format({
        "underline": 0,
        "font_color": "black"
    })

    worksheet.set_column("D:D", None, plain_fmt)
    worksheet.set_column("H:H", None, plain_fmt)

    writer.close()

    return {
        "first_header_sheet": first_header_sheet,
        "rows": combined.shape[0],
        "cols": combined.shape[1],
        "output": out_path,
        "status_msgs": status_msgs,
    }


# ---------- Entry point ----------
def main():
    argv = sys.argv[1:]

    # Input from CLI or hardcoded default
    if len(argv) >= 1:
        input_paths = argv
    else:
        input_paths = [os.path.join(os.getcwd(), "workspace", "output", "Echovita_Family_Links_Tool_Output_Cleaned.xlsx")]

    if not os.path.exists(input_paths[0]):
        print(f"No input found at: {input_paths[0]}")
        sys.exit(1)

    output_path = resolve_single_output_path(input_paths[0])

    print("\n" + "=" * 80)
    print(f"Combining {len(input_paths)} file(s) into ONE output...")
    print(f"Output will be: {output_path}")

    try:
        info = combine_all_files(input_paths, output_path)

        print("✅ Done!")
        print(f"Header taken from FIRST file sheet: {info['first_header_sheet']}")

        for msg in info["status_msgs"]:
            print(msg)

        print(f"Combined size: {info['rows']} rows x {info['cols']} columns")
        print(f"Saved to: {info['output']}")

        sys.exit(0)

    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()