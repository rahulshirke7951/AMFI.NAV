import os
import json
import warnings
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ===================== CONFIG =====================
DEBUG = True   # <<<<<< TURN OFF IN PROD
# ==================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RULES_PATH = os.path.join(BASE_DIR, "scheme_rules.json")

with open(RULES_PATH, "r") as f:
    RULES = json.load(f)

# ===================== RULE VALIDATION =====================

def validate_rules(rules):
    if not rules.get("selection_rules", {}).get("priority_ladder"):
        warnings.warn("Priority ladder is empty – selection may be ambiguous")

    seen = {}
    for t, keys in rules.get("type_rules", {}).items():
        for k in keys:
            if k in seen:
                warnings.warn(f"Keyword '{k}' used in both '{seen[k]}' and '{t}'")
            seen[k] = t

validate_rules(RULES)

# ===================== HELPERS =====================

def normalize_name(s):
    return " ".join(str(s).upper().split())

def clean_text(s):
    for ch in ["-", "–", "—"]:
        s = s.replace(ch, " ")
    return normalize_name(s)

def extract_base_scheme(name):
    s = clean_text(name)
    for term in RULES["base_scheme_remove_terms"]:
        s = s.replace(term, "")
    return normalize_name(s)

def detect_scheme_type(name):
    n = name.upper()
    for t, keys in RULES["type_rules"].items():
        if any(k in n for k in keys):
            return t
    return "Other"

def exclusion_reason(name):
    for k in RULES.get("exclusion_rules", {}).get("contains_any", []):
        if k in name.upper():
            return f"Excluded by rule: contains '{k}'"
    return None

def select_variant(grp, trace):
    for rank, rule in enumerate(RULES["selection_rules"]["priority_ladder"], 1):
        matches = grp[
            grp["Mutual Fund Name"].str.upper().apply(
                lambda x: all(k in x for k in rule)
            )
        ]
        if not matches.empty:
            trace.append(f"Selected by priority {rank}: {rule}")
            return matches.iloc[0], grp.drop(matches.index)
    trace.append("No priority match – defaulted to first scheme")
    return grp.iloc[0], grp.iloc[1:]

# ===================== FLATTEN MERGES =====================

def flatten_merged_cells(file, sheet="NAV Data"):
    wb = openpyxl.load_workbook(file)
    ws = wb[sheet]
    for m in list(ws.merged_cells.ranges):
        r1, c1, r2, c2 = m.min_row, m.min_col, m.max_row, m.max_col
        val = ws.cell(r1, c1).value
        ws.unmerge_cells(str(m))
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = val
    out = file.replace(".xlsx", "_flat.xlsx")
    wb.save(out)
    return out

# ===================== EXTRACT =====================

def extract_nav_data(file):
    flat = flatten_merged_cells(file)
    raw = pd.read_excel(flat, header=None).dropna(how="all")

    hdr = raw[raw.apply(lambda r: r.astype(str).str.contains("NAV Name", case=False).any(), axis=1)].index[0]
    raw.columns = raw.iloc[hdr]
    df = raw.iloc[hdr + 1:][["NAV Name", "Net Asset Value"]]
    df.columns = ["Mutual Fund Name", "NAV"]
    df["NAV"] = pd.to_numeric(df["NAV"], errors="coerce")
    df = df.dropna()

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
    df["Type"] = df["Mutual Fund Name"].apply(detect_scheme_type)
    df["Key"] = df["Mutual Fund Name"].apply(normalize_name)

    final, variant_excl, traces = [], [], []

    for base, grp in df.groupby("Base"):
        trace = [f"Base Scheme: {base}"]
        if len(grp) == 1:
            final.append(grp.iloc[0])
            trace.append("Only one variant – kept")
        else:
            keep, drop = select_variant(grp, trace)
            final.append(keep)
            for _, d in drop.iterrows():
                e = d.to_dict()
                e["Reason"] = "Excluded: non-preferred variant"
                variant_excl.append(e)
        if DEBUG:
            traces.append({"Base Scheme": base, "Decision Trace": " | ".join(trace)})

    return (
        pd.DataFrame(final),
        pd.concat([pd.DataFrame(excluded), pd.DataFrame(variant_excl)], ignore_index=True),
        pd.DataFrame(traces)
    )

# ===================== FORMAT =====================

def format_excel(file):
    wb = openpyxl.load_workbook(file)

    color_map = {
        "Excluded by rule": "FFC7CE",
        "non-preferred": "FFEB9C",
        "missing": "D9D9D9"
    }

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        if "Excluded" in sheet and "Reason" in ws[1]:
            col = [c.value for c in ws[1]].index("Reason") + 1
            for r in ws.iter_rows(min_row=2):
                for k, clr in color_map.items():
                    if k.lower() in str(r[col - 1].value).lower():
                        r[col - 1].fill = PatternFill("solid", fgColor=clr)

    wb.save(file)

# ===================== COMPARE =====================

def compare(latest, past):
    l_df, l_exc, l_trace = extract_nav_data(latest)
    p_df, p_exc, _ = extract_nav_data(past)

    l_df = l_df.rename(columns={"NAV": "Latest NAV"})
    p_df = p_df.rename(columns={"NAV": "Past NAV"})

    m = pd.merge(l_df, p_df[["Key", "Past NAV"]], on="Key", how="left")
    m["Change"] = m["Latest NAV"] - m["Past NAV"]
    m["Change %"] = (m["Change"] / m["Past NAV"] * 100).round(2)

    not_comp = m[m["Past NAV"].isna()].copy()
    not_comp["Reason"] = "Excluded: missing Past NAV"

    comp = m.drop(not_comp.index)

    with pd.ExcelWriter("NAV_Comparison_Result.xlsx", engine="openpyxl") as w:
        comp.to_excel(w, "NAV Comparison", index=False)
        l_exc.to_excel(w, "Excluded_Schemes", index=False)
        not_comp.to_excel(w, "Excluded_Not_Comparable", index=False)
        if DEBUG:
            l_trace.to_excel(w, "Decision_Trace", index=False)

    format_excel("NAV_Comparison_Result.xlsx")

# ===================== RUN =====================

if __name__ == "__main__":
    compare("data/latest_nav.xlsx", "data/past_nav.xlsx")
