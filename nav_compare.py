import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- STEP 1: FLATTEN / UNMERGE CELLS ----------

def flatten_merged_cells(input_file, sheet_name="NAV Data"):
    """
    Open Excel, unmerge all merged cells in the given sheet,
    and fill all cells in those ranges with the top-left value.
    Returns path of a temporary 'flattened' copy.
    """
    wb = openpyxl.load_workbook(input_file)
    ws = wb[sheet_name]

    # Copy merged ranges list first, because we will modify during loop
    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        # e.g. "A1:C1"
        coord = merged_range.coord
        min_row, min_col, max_row, max_col = merged_range.bounds

        # Get value from top-left cell
        top_left_value = ws.cell(row=min_row, column=min_col).value

        # Unmerge the range
        ws.unmerge_cells(coord)

        # Fill all cells in that block with the top-left value
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).value = top_left_value

    # Save to a temporary file
    base, ext = os.path.splitext(input_file)
    flat_file = base + "_flat" + ext
    wb.save(flat_file)
    return flat_file

# ---------- STEP 2: EXTRACT ONLY NAV ROWS ----------

def extract_nav_data(excel_file, sheet_name="NAV Data"):
    """
    1) Flatten merged cells
    2) Read data with pandas
    3) Keep only rows with NAV Name + Net Asset Value
    """
    # First flatten the merged cells
    flat_file = flatten_merged_cells(excel_file, sheet_name=sheet_name)

    # Read without header, then detect header row
    df_raw = pd.read_excel(flat_file, sheet_name=sheet_name, header=None)

    # Drop completely empty rows
    df_raw = df_raw.dropna(how="all")

    # Find row that contains 'NAV Name' (header row)
    header_row_candidates = df_raw.apply(
        lambda r: r.astype(str).str.contains("NAV Name", case=False, na=False).any(),
        axis=1
    )
    header_row_idx = header_row_candidates[header_row_candidates].index[0]

    # Set header
    df_raw.columns = df_raw.iloc[header_row_idx].values
    df = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # We expect these columns in your file
    if "NAV Name" not in df.columns or "Net Asset Value" not in df.columns:
        raise ValueError("Expected columns 'NAV Name' and 'Net Asset Value' not found.")

    # Keep only relevant columns
    df = df[["NAV Name", "Net Asset Value"]].copy()

    # Keep rows that actually have NAV values
    df = df[df["Net Asset Value"].notna()]

    # Rename to consistent names for rest of script
    df = df.rename(columns={
        "NAV Name": "Mutual Fund Name",
        "Net Asset Value": "NAV"
    })

    # Convert NAV to numeric, drop bad rows
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df[df["NAV"].notna()]

    # Drop blank fund names
    df = df[df["Mutual Fund Name"].notna()]

    return df

# ---------- STEP 3: FORMAT OUTPUT ----------

def format_excel_output(df, output_file):
    # Write to Excel
    df.to_excel(output_file, index=False, sheet_name="NAV Comparison")

    wb = openpyxl.load_workbook(output_file)
    ws = wb.active

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

    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(output_file)

# ---------- STEP 4: MAIN COMPARISON ----------

def compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx"):
    latest_df = extract_nav_data(latest_file)
    past_df = extract_nav_data(past_file)

    latest_df = latest_df.rename(columns={"NAV": "Latest NAV"})
    past_df = past_df.rename(columns={"NAV": "Past NAV"})

    merged = pd.merge(
        latest_df,
        past_df[["Mutual Fund Name", "Past NAV"]],
        on="Mutual Fund Name",
        how="inner"
    )

    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = (merged["Change"] / merged["Past NAV"] * 100).round(2)

    merged["Latest NAV"] = merged["Latest NAV"].round(4)
    merged["Past NAV"] = merged["Past NAV"].round(4)
    merged["Change"] = merged["Change"].round(4)

    merged = merged[
        ["Mutual Fund Name", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False)

    format_excel_output(merged, output_file)
    return merged

# For GitHub / local run
if __name__ == "__main__":
    latest_file = "data/latest_nav.xlsx"
    past_file   = "data/past_nav.xlsx"
    compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx")
