import os
import json
import pandas as pd
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, "scheme_rules.json")) as f:
    RULES = json.load(f)

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

def exclusion_reason(name):
    for k in RULES["exclusion_rules"]["contains_any"]:
        if k in name.upper():
            return f"Excluded by rule: contains {k}"
    return None

def select_variant(grp):
    for rule in RULES["selection_rules"]["priority_ladder"]:
        matches = grp[
            grp["Mutual Fund Name"].str.upper().apply(
                lambda x: all(k in x for k in rule)
            )
        ]
        if not matches.empty:
            return matches.iloc[0], grp.drop(matches.index)
    return grp.iloc[0], grp.iloc[1:]

# ======================================================
# FLATTEN MERGED CELLS
# ======================================================

def flatten(file, sheet="NAV Data"):
    wb = openpyxl.load_workbook(file)
    ws = wb[sheet]

    for m in list(ws.merged_cells.ranges):
        val = ws.cell(m.min_row, m.min_col).value
        ws.unmerge_cells(str(m))
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                ws.cell(r, c).value = val

    out = file.replace(".xlsx", "_flat.xlsx")
    wb.save(out)
    return out

# ======================================================
# EXTRACT & CLEAN NAV DATA
# ======================================================

def extract(file):
    flat = flatten(file)
    raw = pd.read_excel(flat, header=None).dropna(how="all")

    header_row = raw[
        raw.apply(lambda r: r.astype(str).str.contains("NAV Name", case=False).any(), axis=1)
    ].index[0]

    raw.columns = raw.iloc[header_row]
    df = raw.iloc[header_row + 1:][["NAV Name", "Net Asset Value"]]
    df.columns = ["Mutual Fund Name", "NAV"]

    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df.dropna(subset=["NAV", "Mutual Fund Name"])

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
    df["Key"] = df["Mutual Fund Name"].apply(normalize)

    final, variant_excluded = [], []

    for _, grp in df.groupby("Base"):
        if len(grp) == 1:
            final.append(grp.iloc[0])
        else:
            keep, drop = select_variant(grp)
            final.append(keep)
            for _, d in drop.iterrows():
                e = d.to_dict()
                e["Reason"] = "Excluded: non-preferred variant"
                variant_excluded.append(e)

    final_df = pd.DataFrame(final)
    excluded_df = pd.concat(
        [pd.DataFrame(excluded), pd.DataFrame(variant_excluded)],
        ignore_index=True
    )

    return final_df, excluded_df, total_raw

# ======================================================
# RECONCILIATION
# ======================================================

def build_reconciliation(lat_raw, past_raw, included, excl_rules, excl_zero, excl_nc):
    rows = []
    for label, raw in [("Latest", lat_raw), ("Past", past_raw)]:
        rows.extend([
            [label, "Total Raw", raw],
            [label, "Included in NAV Comparison", len(included)],
            [label, "Excluded – Scheme Rules", len(excl_rules)],
            [label, "Excluded – Zero NAV", len(excl_zero)],
            [label, "Excluded – Not Comparable", len(excl_nc)],
            [label, "Total Check", raw]
        ])
    return pd.DataFrame(rows, columns=["File Type", "Category", "Count"])

# ======================================================
# MAIN COMPARISON
# ======================================================

def run(latest, past):
    l_df, l_exc, l_raw = extract(latest)
    p_df, p_exc, p_raw = extract(past)

    l_df = l_df.rename(columns={"NAV": "Latest NAV"})
    p_df = p_df.rename(columns={"NAV": "Past NAV"})

    merged = pd.merge(
        l_df,
        p_df[["Key", "Past NAV"]],
        on="Key",
        how="left"
    )

    # Missing past NAV
    excl_nc = merged[merged["Past NAV"].isna()].copy()
    excl_nc["Reason"] = "Excluded: missing Past NAV"
    merged = merged.drop(excl_nc.index)

    # Zero NAV
    excl_zero = merged[(merged["Latest NAV"] == 0) | (merged["Past NAV"] == 0)].copy()
    excl_zero["Reason"] = "Excluded: zero NAV"
    merged = merged.drop(excl_zero.index)

    merged["Change"] = merged["Latest NAV"] - merged["Past NAV"]
    merged["Change %"] = (merged["Change"] / merged["Past NAV"] * 100).round(2)

    nav_comparison = merged[
        ["Mutual Fund Name", "Latest NAV", "Past NAV", "Change", "Change %"]
    ].sort_values("Change %", ascending=False)

    reconciliation = build_reconciliation(
        l_raw, p_raw, nav_comparison, l_exc, excl_zero, excl_nc
    )

    with pd.ExcelWriter("NAV_Comparison_Result.xlsx", engine="openpyxl") as w:
        nav_comparison.to_excel(w, "NAV Comparison", index=False)
        l_exc.to_excel(w, "Excluded_Scheme_Rules", index=False)
        excl_zero.to_excel(w, "Excluded_Zero_NAV", index=False)
        excl_nc.to_excel(w, "Excluded_Not_Comparable", index=False)
        reconciliation.to_excel(w, "Reconciliation", index=False)

# ======================================================
# ENTRY POINT
# ======================================================

if __name__ == "__main__":
    run("data/latest_nav.xlsx", "data/past_nav.xlsx")
