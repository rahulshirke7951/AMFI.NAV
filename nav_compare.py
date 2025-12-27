import os
import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DEBUG = True

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, "scheme_rules.json")) as f:
    RULES = json.load(f)

with open(os.path.join(BASE_DIR, "formatting_rules.json")) as f:
    FORMAT = json.load(f)

# ======================================================
# HELPERS
# ======================================================

def normalize(s):
    return " ".join(str(s).upper().split())

def clean_text(s):
    for ch in ["-", "–", "—"]:
        s = s.replace(ch, " ")
    return normalize(s)

def extract_base_scheme(name):
    s = clean_text(name)
    for t in RULES["base_scheme_remove_terms"]:
        s = s.replace(t, "")
    return normalize(s)

def detect_type(name):
    n = name.upper()
    for t, keys in RULES["type_rules"].items():
        if any(k in n for k in keys):
            return t
    return "Other"

def exclusion_reason(name):
    for k in RULES["exclusion_rules"]["contains_any"]:
        if k in name.upper():
            return f"Excluded by rule: contains {k}"
    return None

def select_variant(grp):
    for rule in RULES["selection_rules"]["priority_ladder"]:
        m = grp[grp["Mutual Fund Name"].str.upper().apply(lambda x: all(k in x for k in rule))]
        if not m.empty:
            return m.iloc[0], grp.drop(m.index)
    return grp.iloc[0], grp.iloc[1:]

# ======================================================
# FLATTEN MERGES
# ======================================================

def flatten(file, sheet="NAV Data"):
    wb = openpyxl.load_workbook(file)
    ws = wb[sheet]
    for m in list(ws.merged_cells.ranges):
        v = ws.cell(m.min_row, m.min_col).value
        ws.unmerge_cells(str(m))
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                ws.cell(r, c).value = v
    out = file.replace(".xlsx", "_flat.xlsx")
    wb.save(out)
    return out

# ======================================================
# EXTRACT
# ======================================================

def extract(file):
    raw_flat = flatten(file)
    raw = pd.read_excel(raw_flat, header=None).dropna(how="all")

    hdr = raw[raw.apply(lambda r: r.astype(str).str.contains("NAV Name", case=False).any(), axis=1)].index[0]
    raw.columns = raw.iloc[hdr]
    df = raw.iloc[hdr + 1:][["NAV Name", "Net Asset Value"]]
    df.columns = ["Mutual Fund Name", "NAV"]
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df.dropna()

    total_raw = len(df)

    included, excluded = [], []

    for _, r in df.iterrows():
        reason = exclusion_reason(r["Mutual Fund Name"])
        if reason:
            e = r.to_dict()
            e["Reason"] = reason
            excluded.append(e)
        else:
            included.append(r)

    df = pd.DataFrame(included)
    df["Base"] = df["Mutual Fund Name"].apply(extract_base_scheme)
    df["Type"] = df["Mutual Fund Name"].apply(detect_type)
    df["Key"] = df["Mutual Fund Name"].apply(normalize)

    final, var_excl = [], []

    for _, grp in df.groupby("Base"):
        if len(grp) == 1:
            final.append(grp.iloc[0])
        else:
            keep, drop = select_variant(grp)
            final.append(keep)
            for _, d in drop.iterrows():
                e = d.to_dict()
                e["Reason"] = "Excluded: non-preferred variant"
                var_excl.append(e)

    final_df = pd.DataFrame(final)
    excluded_df = pd.concat([pd.DataFrame(excluded), pd.DataFrame(var_excl)], ignore_index=True)

    return final_df, excluded_df, total_raw

# ======================================================
# FORMAT NAV COMPARISON
# ======================================================

def format_nav_sheet(wb):
    ws = wb["NAV Comparison"]
    nav_fmt = FORMAT["nav_comparison"]

    border = Border(*(Side(style="thin"),) * 4)
    hdr = nav_fmt["header"]

    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor=hdr["fill_color"])
        c.font = Font(bold=hdr["bold"], color=hdr["font_color"])
        c.alignment = Alignment(horizontal=hdr["align"])
        c.border = border

    widths = nav_fmt["column_widths"]
    headers = [c.value for c in ws[1]]

    for i, h in enumerate(headers, 1):
        if h in widths:
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths[h]

    pos = PatternFill("solid", fgColor=nav_fmt["change_percent_colors"]["positive"])
    neg = PatternFill("solid", fgColor=nav_fmt["change_percent_colors"]["negative"])

    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.border = border
        cp = r[headers.index("Change %")]
        if isinstance(cp.value, (int, float)):
            cp.fill = pos if cp.value > 0 else neg

# ======================================================
# RECONCILIATION
# ======================================================

def reconciliation(lat_raw, past_raw, comp, excl_rules, excl_nc):
    rows = []
    for label, raw, inc, exc_r, exc_nc in [
        ("Latest", lat_raw, len(comp), len(excl_rules), len(excl_nc)),
        ("Past", past_raw, len(comp), len(excl_rules), len(excl_nc))
    ]:
        rows.extend([
            [label, "Total Raw", raw],
            [label, "Included in NAV Comparison", inc],
            [label, "Excluded – Scheme Rules", exc_r],
            [label, "Excluded – Not Comparable", exc_nc],
            [label, "Total Check", raw]
        ])
    return pd.DataFrame(rows, columns=["File Type", "Category", "Count"])

# ======================================================
# MAIN
# ======================================================

def run(latest, past):
    l_df, l_exc, l_raw = extract(latest)
    p_df, p_exc, p_raw = extract(past)

    l_df = l_df.rename(columns={"NAV": "Latest NAV"})
    p_df = p_df.rename(columns={"NAV": "Past NAV"})

    m = pd.merge(l_df, p_df[["Key", "Past NAV"]], on="Key", how="left")
    m["Change"] = m["Latest NAV"] - m["Past NAV"]
    m["Change %"] = (m["Change"] / m["Past NAV"] * 100).round(2)

    non_comp = m[m["Past NAV"].isna()].copy()
    non_comp["Reason"] = "Excluded: missing Past NAV"

    comp = m.drop(non_comp.index)[
        ["Mutual Fund Name", "Type", "Latest NAV", "Past NAV", "Change", "Change %"]
    ]

    recon = reconciliation(l_raw, p_raw, comp, l_exc, non_comp)

    with pd.ExcelWriter("NAV_Comparison_Result.xlsx", engine="openpyxl") as w:
        comp.to_excel(w, "NAV Comparison", index=False)
        l_exc.to_excel(w, "Excluded_Scheme_Rules", index=False)
        non_comp.to_excel(w, "Excluded_Not_Comparable", index=False)
        recon.to_excel(w, "Reconciliation", index=False)

    wb = openpyxl.load_workbook("NAV_Comparison_Result.xlsx")
    format_nav_sheet(wb)
    wb.save("NAV_Comparison_Result.xlsx")

# ======================================================
# ENTRY
# ======================================================

if __name__ == "__main__":
    run("data/latest_nav.xlsx", "data/past_nav.xlsx")
