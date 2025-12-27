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
# ADVISORY
# ======================================================

def build_advisory(comp_df):
    rows = []

    if comp_df.empty:
        return pd.DataFrame(columns=["Section", "Message"])

    strong = comp_df[comp_df["Change %"] >= 5]
    weak = comp_df[comp_df["Change %"] <= -5]
    stable = comp_df[comp_df["Change %"].abs() < 1]

    rows.append(["Strong Performers",
                 f"{len(strong)} schemes gained more than 5% in the period."])

    rows.append(["Area of Attention",
                 f"{len(weak)} schemes declined more than -5% and may require review."])

    rows.append(["Stable Schemes",
                 f"{len(stable)} schemes moved within ±1%, indicating stability."])

    cat = comp_df.groupby("Type")["Change %"].mean().round(2)
    for t, v in cat.items():
        rows.append(["Category Insight",
                     f"{t} schemes averaged {v}% change in NAV."])

    return pd.DataFrame(rows, columns=["Section", "Message"])

# ======================================================
# RECONCILIATION
# ======================================================

def reconciliation(lat_raw, past_raw, comp, excl_rules, excl_zero, excl_nc):
    rows = []
    for label, raw in [("Latest", lat_raw), ("Past", past_raw)]:
        rows.extend([
            [label, "Total Raw", raw],
            [label, "Included in NAV Comparison", len(comp)],
            [label, "Excluded – Scheme Rules", len(excl_rules)],
            [label, "Excluded – Zero NAV", len(excl_zero)],
            [label, "Excluded – Not Comparable", len(excl_nc)],
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

    merged = pd.merge(l_df, p_df[["Key", "Past NAV"]], on="Key", how="left")

    # Missing NAV
    excl_nc = merged[merged["Past NAV"].isna()].copy()
    excl_nc["Reason"] = "Excluded: missing Past NAV"

    merged = merged.drop(excl_nc.index)

    # Zero NAV
    excl_zero = merged[(merged["Latest NAV"] == 0) | (merged["Past NAV"] == 0)].copy()
    excl_zero["Reason"] = "Excluded: Zero NAV"

    merged = merged.drop(excl_zero.index)

    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = (merged["Change"] / merged["Past NAV"] * 100).round(2)

    comp = merged[
        ["Mutual Fund Name", "Type", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False)

    recon = reconciliation(l_raw, p_raw, comp, l_exc, excl_zero, excl_nc)
    advisory = build_advisory(comp)

    with pd.ExcelWriter("NAV_Comparison_Result.xlsx", engine="openpyxl") as w:
        comp.to_excel(w, "NAV Comparison", index=False)
        l_exc.to_excel(w, "Excluded_Scheme_Rules", index=False)
        excl_zero.to_excel(w, "Excluded_Zero_NAV", index=False)
        excl_nc.to_excel(w, "Excluded_Not_Comparable", index=False)
        recon.to_excel(w, "Reconciliation", index=False)
        advisory.to_excel(w, "Advisory", index=False)

# ======================================================
# ENTRY
# ======================================================

if __name__ == "__main__":
    run("data/latest_nav.xlsx", "data/past_nav.xlsx")
