import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- HELPER: NORMALIZE SCHEME NAMES FOR MATCHING ----------

def normalize_name(s):
    """
    Normalize scheme / fund names so minor differences don't break matching.
    - Uppercase
    - Strip spaces
    - Collapse multiple spaces
    """
    if pd.isna(s):
        return None
    s = str(s).upper().strip()
    # Replace multiple spaces with single
    while "  " in s:
        s = s.replace("  ", " ")
    return s


# ---------- STEP 1: FLATTEN / UNMERGE MERGED CELLS ----------

def flatten_merged_cells(input_file, sheet_name="NAV Data"):
    """
    Open Excel, unmerge all merged cells in the given sheet,
    and fill all cells in those ranges with the top-left value.
    Returns path of a temporary 'flattened' copy.
    """
    wb = openpyxl.load_workbook(input_file)
    ws = wb[sheet_name]

    # Copy merged ranges first, because we'll modify them in the loop
    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        # Bounds of the merged block
        min_row = merged_range.min_row
        min_col = merged_range.min_col
        max_row = merged_range.max_row
        max_col = merged_range.max_col

        # Value from top-left cell
        top_left_value = ws.cell(row=min_row, column=min_col).value

        # Collect all cell coordinates BEFORE unmerging
        cells_to_fill = [
            (r, c)
            for r in range(min_row, max_row + 1)
            for c in range(min_col, max_col + 1)
        ]

        # Unmerge this range
        ws.unmerge_cells(range_string=str(merged_range))

        # Now all cells are normal (not MergedCell), safe to write
        for r, c in cells_to_fill:
            ws.cell(row=r, column=c).value = top_left_value

    # Save to a temporary "flat" file
    base, ext = os.path.splitext(input_file)
    flat_file = base + "_flat" + ext
    wb.save(flat_file)
    return flat_file


# ---------- STEP 2: EXTRACT ONLY NAV ROWS ----------

def extract_nav_data(excel_file, sheet_name="NAV Data"):
    """
    1) Flatten merged cells
    2) Read with pandas
    3) Detect header row (contains 'NAV Name')
    4) Keep only rows with NAV values
    Returns: DataFrame with columns ['Mutual Fund Name', 'NAV', 'Key']
    """
    # First flatten merged cells into a temp file
    flat_file = flatten_merged_cells(excel_file, sheet_name=sheet_name)

    # Read raw sheet without header
    df_raw = pd.read_excel(flat_file, sheet_name=sheet_name, header=None)

    # Drop completely empty rows
    df_raw = df_raw.dropna(how="all")

    # Find index of row that contains 'NAV Name'
    header_row_mask = df_raw.apply(
        lambda r: r.astype(str).str.contains("NAV Name", case=False, na=False).any(),
        axis=1
    )
    header_rows = header_row_mask[header_row_mask].index.tolist()
    if not header_rows:
        raise ValueError(f"'NAV Name' header not found in file: {excel_file}")
    header_row_idx = header_rows[0]

    # Use that row as header
    df_raw.columns = df_raw.iloc[header_row_idx].values
    df = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Expect standard columns
    if "NAV Name" not in df.columns or "Net Asset Value" not in df.columns:
        raise ValueError(
            f"Expected columns 'NAV Name' and 'Net Asset Value' not found in file: {excel_file}"
        )

    # Keep only relevant columns
    df = df[["NAV Name", "Net Asset Value"]].copy()

    # Keep rows with actual NAV values
    df = df[df["Net Asset Value"].notna()]

    # Rename for consistency
    df = df.rename(columns={
        "NAV Name": "Mutual Fund Name",
        "Net Asset Value": "NAV"
    })

    # Convert NAV to numeric, drop invalid rows
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df[df["NAV"].notna()]

    # Drop blank fund names
    df = df[df["Mutual Fund Name"].notna()]

    # Normalize name for matching
    df["Key"] = df["Mutual Fund Name"].apply(normalize_name)

    # Drop rows where key could not be generated
    df = df[df["Key"].notna()]

    # Drop duplicate keys (if any)
    df = df.drop_duplicates(subset=["Key"])

    return df


# ---------- STEP 3: FORMAT THE OUTPUT EXCEL (ONLY MAIN SHEET) ----------

def format_excel_output(output_file, sheet_name="NAV Comparison"):
    wb = openpyxl.load_workbook(output_file)
    if sheet_name not in wb.sheetnames:
        wb.save(output_file)
        return

    ws = wb[sheet_name]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Header styling
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Data rows styling
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for col_idx, cell in enumerate(row, start=1):
            cell.border = border
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")

            # Change % column = 5
            if col_idx == 5:
                try:
                    val = float(cell.value)
                    if val > 0:
                        cell.fill = green_fill
                    elif val < 0:
                        cell.fill = red_fill
                    cell.font = Font(bold=True)
                except (TypeError, ValueError):
                    pass

    # Column widths
    ws.column_dimensions["A"].width = 60
    for col in ["B", "C", "D", "E"]:
        ws.column_dimensions[col].width = 15

    # Auto filter
    ws.auto_filter.ref = f"A1:E{ws.max_row}"

    wb.save(output_file)


# ---------- STEP 4: MAIN COMPARISON LOGIC + OUTLIER LISTS ----------

def compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx"):
    # Extract clean data from both files
    latest_df = extract_nav_data(latest_file)
    past_df   = extract_nav_data(past_file)

    # Rename NAV columns for clarity
    latest_df = latest_df.rename(columns={"NAV": "Latest NAV"})
    past_df   = past_df.rename(columns={"NAV": "Past NAV"})

    # LEFT JOIN on normalized key so all latest schemes are kept
    merged = pd.merge(
        latest_df[["Key", "Mutual Fund Name", "Latest NAV"]],
        past_df[["Key", "Past NAV"]],
        on="Key",
        how="left"
    )

    # Compute Change and Change %
    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = (merged["Change"] / merged["Past NAV"] * 100)

    # Rounding
    merged["Latest NAV"] = merged["Latest NAV"].round(4)
    merged["Past NAV"]   = merged["Past NAV"].round(4)
    merged["Change"]     = merged["Change"].round(4)
    merged["Change %"]   = merged["Change %"].round(2)

    # Final column order & sorting (missing % at bottom)
    merged = merged[
        ["Mutual Fund Name", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False, na_position="last")

    # ---- Outlier lists ----
    # Schemes present in latest but NOT in past
    missing_in_past = latest_df[~latest_df["Key"].isin(past_df["Key"])][
        ["Mutual Fund Name", "Latest NAV"]
    ].sort_values("Mutual Fund Name")

    # Schemes present in past but NOT in latest
    missing_in_latest = past_df[~past_df["Key"].isin(latest_df["Key"])][
        ["Mutual Fund Name", "Past NAV"]
    ].sort_values("Mutual Fund Name")

    # ---- Write all sheets ----
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Main comparison
        merged.to_excel(writer, sheet_name="NAV Comparison", index=False)

        # Outlier sheets (only if non-empty)
        if not missing_in_past.empty:
            missing_in_past.to_excel(writer, sheet_name="Missing_in_Past", index=False)

        if not missing_in_latest.empty:
            missing_in_latest.to_excel(writer, sheet_name="Missing_in_Latest", index=False)

    # Apply formatting to main sheet
    format_excel_output(output_file, sheet_name="NAV Comparison")

    return merged


# ---------- STEP 5: ENTRY POINT (for GitHub Action / local run) ----------

if __name__ == "__main__":
    # These are the files your GitHub repo / local setup should have
    latest_file = "data/latest_nav.xlsx"
    past_file   = "data/past_nav.xlsx"

    compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx")
