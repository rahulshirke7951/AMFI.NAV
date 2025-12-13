import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- HELPER: NORMALIZE SCHEME NAMES ----------

def normalize_name(s):
    if pd.isna(s):
        return None
    s = str(s).upper().strip()
    while "  " in s:
        s = s.replace("  ", " ")
    return s


# ---------- HELPER: BASE SCHEME & PRIORITY ----------

def extract_base_scheme(name):
    if pd.isna(name):
        return None

    s = str(name).upper()

    remove_terms = [
        "DIRECT PLAN", "REGULAR PLAN",
        "GROWTH", "IDCW"
    ]

    for term in remove_terms:
        s = s.replace(term, "")

    s = s.replace(" - ", " ").strip()
    while "  " in s:
        s = s.replace("  ", " ")

    return s


def variant_priority(name):
    n = name.upper()

    if "REGULAR PLAN" in n and "GROWTH" in n:
        return 1
    if "REGULAR PLAN" in n:
        return 2
    if "DIRECT PLAN" in n and "GROWTH" in n:
        return 3
    return 99


# ---------- STEP 1: FLATTEN MERGED CELLS ----------

def flatten_merged_cells(input_file, sheet_name="NAV Data"):
    wb = openpyxl.load_workbook(input_file)
    ws = wb[sheet_name]

    merged_ranges = list(ws.merged_cells.ranges)

    for merged_range in merged_ranges:
        min_row, min_col = merged_range.min_row, merged_range.min_col
        max_row, max_col = merged_range.max_row, merged_range.max_col

        value = ws.cell(row=min_row, column=min_col).value
        cells = [(r, c) for r in range(min_row, max_row + 1)
                        for c in range(min_col, max_col + 1)]

        ws.unmerge_cells(str(merged_range))

        for r, c in cells:
            ws.cell(row=r, column=c).value = value

    base, ext = os.path.splitext(input_file)
    flat_file = base + "_flat" + ext
    wb.save(flat_file)
    return flat_file


# ---------- STEP 2: EXTRACT & FILTER NAV DATA ----------

def extract_nav_data(excel_file, sheet_name="NAV Data"):
    flat_file = flatten_merged_cells(excel_file, sheet_name)

    df_raw = pd.read_excel(flat_file, sheet_name=sheet_name, header=None)
    df_raw = df_raw.dropna(how="all")

    header_row = df_raw[
        df_raw.apply(lambda r: r.astype(str).str.contains("NAV Name", case=False).any(), axis=1)
    ].index[0]

    df_raw.columns = df_raw.iloc[header_row].values
    df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = df.columns.astype(str).str.strip()

    df = df[["NAV Name", "Net Asset Value"]]
    df = df.rename(columns={
        "NAV Name": "Mutual Fund Name",
        "Net Asset Value": "NAV"
    })

    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df.dropna(subset=["NAV", "Mutual Fund Name"])

    # ---------- VARIANT FILTERING ----------
    df["Base Scheme"] = df["Mutual Fund Name"].apply(extract_base_scheme)
    df["Priority"] = df["Mutual Fund Name"].apply(variant_priority)
    df["Key"] = df["Mutual Fund Name"].apply(normalize_name)

    final_rows = []
    excluded_rows = []

    for base, grp in df.groupby("Base Scheme"):
        if len(grp) == 1:
            final_rows.append(grp.iloc[0])
        else:
            grp = grp.sort_values("Priority")
            final_rows.append(grp.iloc[0])
            excluded_rows.append(grp.iloc[1:])

    final_df = pd.DataFrame(final_rows).reset_index(drop=True)
    excluded_df = (
        pd.concat(excluded_rows).reset_index(drop=True)
        if excluded_rows else
        pd.DataFrame(columns=df.columns)
    )

    final_df = final_df.drop(columns=["Base Scheme", "Priority"])
    excluded_df = excluded_df.drop(columns=["Base Scheme", "Priority"])

    return final_df, excluded_df


# ---------- STEP 3: FORMAT OUTPUT ----------

def format_excel_output(output_file, sheet_name="NAV Comparison"):
    wb = openpyxl.load_workbook(output_file)
    ws = wb[sheet_name]

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    border = Border(*(Side(style="thin"),) * 4)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")

    for row in ws.iter_rows(min_row=2):
        for i, cell in enumerate(row, 1):
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if i == 1 else "right")
            if i == 5:
                try:
                    if cell.value > 0:
                        cell.fill = green
                    elif cell.value < 0:
                        cell.fill = red
                    cell.font = Font(bold=True)
                except:
                    pass

    ws.column_dimensions["A"].width = 60
    for c in "BCDE":
        ws.column_dimensions[c].width = 15

    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(output_file)


# ---------- STEP 4: COMPARE ----------

def compare_nav_files(latest_file, past_file, output_file="NAV_Comparison_Result.xlsx"):
    latest_df, latest_excl = extract_nav_data(latest_file)
    past_df, past_excl = extract_nav_data(past_file)

    latest_df = latest_df.rename(columns={"NAV": "Latest NAV"})
    past_df = past_df.rename(columns={"NAV": "Past NAV"})

    merged = pd.merge(
        latest_df[["Key", "Mutual Fund Name", "Latest NAV"]],
        past_df[["Key", "Past NAV"]],
        on="Key",
        how="left"
    )

    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = merged["Change"] / merged["Past NAV"] * 100

    merged = merged.round({
        "Latest NAV": 4,
        "Past NAV": 4,
        "Change": 4,
        "Change %": 2
    })

    merged = merged[
        ["Mutual Fund Name", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False, na_position="last")

    missing_in_past = latest_df[~latest_df["Key"].isin(past_df["Key"])][
        ["Mutual Fund Name", "Latest NAV"]
    ]

    missing_in_latest = past_df[~past_df["Key"].isin(latest_df["Key"])][
        ["Mutual Fund Name", "Past NAV"]
    ]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="NAV Comparison", index=False)
        missing_in_past.to_excel(writer, sheet_name="Missing_in_Past", index=False)
        missing_in_latest.to_excel(writer, sheet_name="Missing_in_Latest", index=False)
        latest_excl.to_excel(writer, sheet_name="Excluded_Latest_Variants", index=False)
        past_excl.to_excel(writer, sheet_name="Excluded_Past_Variants", index=False)

    format_excel_output(output_file)
    return merged


# ---------- ENTRY POINT ----------

if __name__ == "__main__":
    compare_nav_files(
        latest_file="data/latest_nav.xlsx",
        past_file="data/past_nav.xlsx",
        output_file="NAV_Comparison_Result.xlsx"
    )
