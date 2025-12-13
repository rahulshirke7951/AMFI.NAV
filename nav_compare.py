import os
import json
import warnings
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =========================================================
# LOAD JSON RULES (PATH SAFE)
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RULES_PATH = os.path.join(BASE_DIR, "scheme_rules.json")

with open(RULES_PATH, "r") as f:
    RULES = json.load(f)

# =========================================================
# RULE VALIDATION (WARNINGS ONLY)
# =========================================================

def validate_rules(rules):
    warnings_list = []

    if not rules.get("base_scheme_remove_terms"):
        warnings_list.append("base_scheme_remove_terms is empty")

    sel = rules.get("selection_rules", {}).get("preferred_variant", {})
    if not sel.get("must_contain_all"):
        warnings_list.append("preferred_variant.must_contain_all is empty")

    type_keywords = {}
    for t, keys in rules.get("type_rules", {}).items():
        for k in keys:
            if k in type_keywords:
                warnings_list.append(
                    f"Keyword '{k}' used in both '{type_keywords[k]}' and '{t}'"
                )
            type_keywords[k] = t

    for w in warnings_list:
        warnings.warn(f"[RULE WARNING] {w}")

validate_rules(RULES)

# =========================================================
# RULE ENGINE HELPERS
# =========================================================

def normalize_name(s):
    if pd.isna(s):
        return None
    s = str(s).upper().strip()
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def extract_base_scheme(name):
    s = name.upper()
    for ch in ["-", "–", "—"]:
        s = s.replace(ch, " ")

    for term in RULES["base_scheme_remove_terms"]:
        s = s.replace(term, "")

    while "  " in s:
        s = s.replace("  ", " ")

    return s.strip()


def is_excluded(name):
    n = name.upper()
    for k in RULES.get("exclusion_rules", {}).get("contains_any", []):
        if k in n:
            return True, f"Excluded by rule: contains '{k}'"
    return False, None


def is_preferred_variant(name):
    n = name.upper()
    must_have = RULES["selection_rules"]["preferred_variant"]["must_contain_all"]
    return all(k in n for k in must_have)


def detect_scheme_type(name):
    n = name.upper()
    for scheme_type, keywords in RULES["type_rules"].items():
        if any(k in n for k in keywords):
            return scheme_type
    return "Other"


# =========================================================
# FLATTEN MERGED CELLS
# =========================================================

def flatten_merged_cells(input_file, sheet_name="NAV Data"):
    wb = openpyxl.load_workbook(input_file)
    ws = wb[sheet_name]

    for m in list(ws.merged_cells.ranges):
        r1, c1, r2, c2 = m.min_row, m.min_col, m.max_row, m.max_col
        val = ws.cell(r1, c1).value
        ws.unmerge_cells(str(m))
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = val

    base, ext = os.path.splitext(input_file)
    out = base + "_flat" + ext
    wb.save(out)
    return out


# =========================================================
# EXTRACT & APPLY SCHEME RULES
# =========================================================

def extract_nav_data(excel_file, sheet_name="NAV Data"):
    flat = flatten_merged_cells(excel_file, sheet_name)

    df_raw = pd.read_excel(flat, sheet_name=sheet_name, header=None).dropna(how="all")

    header_row = df_raw[
        df_raw.apply(
            lambda r: r.astype(str).str.contains("NAV Name", case=False).any(),
            axis=1
        )
    ].index[0]

    df_raw.columns = df_raw.iloc[header_row]
    df = df_raw.iloc[header_row + 1:].reset_index(drop=True)

    df = df[["NAV Name", "Net Asset Value"]]
    df.columns = ["Mutual Fund Name", "NAV"]
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df.dropna(subset=["NAV", "Mutual Fund Name"])

    included_rows = []
    excluded_rows = []

    for _, row in df.iterrows():
        excluded, reason = is_excluded(row["Mutual Fund Name"])
        if excluded:
            r = row.to_dict()
            r["Reason"] = reason
            excluded_rows.append(r)
        else:
            included_rows.append(row)

    df = pd.DataFrame(included_rows)

    df["Base Scheme"] = df["Mutual Fund Name"].apply(extract_base_scheme)
    df["Type"] = df["Mutual Fund Name"].apply(detect_scheme_type)
    df["Key"] = df["Mutual Fund Name"].apply(normalize_name)

    final_rows = []
    variant_excluded = []

    for base, grp in df.groupby("Base Scheme"):
        if len(grp) == 1:
            final_rows.append(grp.iloc[0])
        else:
            preferred = grp[grp["Mutual Fund Name"].apply(is_preferred_variant)]
            if not preferred.empty:
                final_rows.append(preferred.iloc[0])
                for _, r in grp.drop(preferred.index).iterrows():
                    d = r.to_dict()
                    d["Reason"] = "Excluded: non-preferred variant"
                    variant_excluded.append(d)
            else:
                final_rows.append(grp.iloc[0])
                for _, r in grp.iloc[1:].iterrows():
                    d = r.to_dict()
                    d["Reason"] = "Excluded: non-preferred variant"
                    variant_excluded.append(d)

    final_df = pd.DataFrame(final_rows).reset_index(drop=True)

    excluded_df = pd.concat(
        [
            pd.DataFrame(excluded_rows),
            pd.DataFrame(variant_excluded)
        ],
        ignore_index=True
    )

    return (
        final_df.drop(columns=["Base Scheme"]),
        excluded_df.drop(columns=["Base Scheme"], errors="ignore")
    )


# =========================================================
# FORMAT OUTPUT
# =========================================================

def format_excel_output(file):
    wb = openpyxl.load_workbook(file)
    ws = wb["NAV Comparison"]

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    border = Border(*(Side(style="thin"),) * 4)

    for c in ws[1]:
        c.fill = header_fill
        c.font = header_font
        c.border = border
        c.alignment = Alignment(horizontal="center")

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")

    for r in ws.iter_rows(min_row=2):
        for i, c in enumerate(r, 1):
            c.border = border
            c.alignment = Alignment(horizontal="left" if i == 1 else "right")
            if i == 6 and isinstance(c.value, (int, float)):
                c.fill = green if c.value > 0 else red
                c.font = Font(bold=True)

    wb.save(file)


# =========================================================
# COMPARE NAV FILES
# =========================================================

def compare_nav_files(latest, past, output="NAV_Comparison_Result.xlsx"):
    l_df, l_exc = extract_nav_data(latest)
    p_df, p_exc = extract_nav_data(past)

    l_df = l_df.rename(columns={"NAV": "Latest NAV"})
    p_df = p_df.rename(columns={"NAV": "Past NAV"})

    merged = pd.merge(
        l_df,
        p_df[["Key", "Past NAV"]],
        on="Key",
        how="left"
    )

    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = (merged["Change"] / merged["Past NAV"] * 100).round(2)

    non_comparable = merged[
        merged["Latest NAV"].isna() | merged["Past NAV"].isna()
    ].copy()

    non_comparable["Reason"] = "Excluded: missing Latest or Past NAV"

    comparable = merged.drop(non_comparable.index).copy()

    comparable = comparable[
        ["Mutual Fund Name", "Type", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False)

    with pd.ExcelWriter(output, engine="openpyxl") as w:
        comparable.to_excel(w, "NAV Comparison", index=False)
        l_exc.to_excel(w, "Excluded_Scheme_Rules", index=False)
        p_exc.to_excel(w, "Excluded_Scheme_Rules_Past", index=False)
        if not non_comparable.empty:
            non_comparable.to_excel(w, "Excluded_Not_Comparable", index=False)

    format_excel_output(output)


# =========================================================
# ENTRY POINT
# =========================================================

if __name__ == "__main__":
    compare_nav_files(
        "data/latest_nav.xlsx",
        "data/past_nav.xlsx"
    )
