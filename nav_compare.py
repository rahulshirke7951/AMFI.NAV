import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def extract_nav_data(excel_file, sheet_name="NAV Data"):
    # Read without header, then set row 1 as header (typical AMFI layout)
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    df.columns = df.iloc[1].values
    df = df.iloc[2:].reset_index(drop=True)

    # Keep only rows with actual NAV values
    df = df[["NAV Name", "Net Asset Value"]].copy()
    df = df[df["Net Asset Value"].notna()]

    df = df.rename(columns={
        "NAV Name": "Mutual Fund Name",
        "Net Asset Value": "NAV"
    })
    df["NAV"] = df["NAV"].astype(float)

    return df

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

if __name__ == "__main__":
    latest_file = "data/latest_nav.xlsx"
    past_file   = "data/past_nav.xlsx"
    compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx")
