import io
import re
import math
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# ACTIVE UZIO EEs vs ADP Comparison (Mapped Columns Only)
# Input Excel must include 3 tabs:
#   - UZIO Data
#   - ADP Data
#   - Mapping Sheet
#
# Logic:
#   - Consider ONLY ACTIVE employees from UZIO Data (employment_status == 'ACTIVE')
#   - For those active EEs, find match in ADP Data (key: employee_id -> Associate ID)
#   - Compare ONLY columns present in Mapping Sheet
#   - UZIO cents -> dollars for specific UZIO fields (divide by 100) before comparing
#   - UZIO TRUE/FALSE matches ADP YES/NO
#   - ADP filing-status text matches specific UZIO enum codes (treat as OK)
#
# Output Excel tabs:
#   - Summary_By_Field
#   - Mismatches_Detail
# ============================================================

APP_TITLE = "UZIO vs ADP (Active Employees Only) — Mismatch Finder"
OUTPUT_FILENAME = "UZIO_ADP_mismatches_active_only.xlsx"

# Expected (but we also match case-insensitively and by 'contains')
UZIO_SHEET_PREFERRED = "UZIO Data"
ADP_SHEET_PREFERRED = "ADP Data"
MAP_SHEET_PREFERRED = "Mapping Sheet"

# Mapping sheet headers (we also auto-detect if these differ)
UZ_MAP_HEADER_PREFERRED = "Uzio Columns"
ADP_MAP_HEADER_PREFERRED = "ADP Columns"

# Keys (used in your sample file)
UZIO_KEY_PREFERRED = "employee_id"
ADP_KEY_PREFERRED = "Associate ID"

# Fields where UZIO stores cents and ADP stores dollars
CENTS_FIELDS = {
    "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
    "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",
    "FIT_DEDUCTIONS_OVER_STANDARD",
    "FIT_OTHER_INCOME",
    "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
}

# Treat blanks as 0 for these numeric fields (helps when ADP stores 0 as blank)
ZERO_IF_BLANK_FIELDS = set(CENTS_FIELDS) | {"SIT_TOTAL_ALLOWANCES"}

# ADP value text -> UZIO value code equivalence (treat as match)
SPECIAL_ADP_UZIO_EQUIVALENCE = {
    ("head of household", "federal_head_of_household"),
    ("single or married filing separately", "federal_single_or_married"),
    ("married filing jointly or qualifying surviving spouse", "federal_married_jointly"),
    ("single", "md_single"),
    ("married", "md_married"),
}

BOOL_TOKEN_MAP = {
    "true": "yes",
    "false": "no",
    "yes": "yes",
    "no": "no",
    "y": "yes",
    "n": "no",
    "1": "yes",
    "0": "no",
}

# ---------------- UI: Hide sidebar/chrome (optional) ----------------
st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
st.markdown(
    """
    <style>
      [data-testid="stSidebar"] { display: none !important; }
      [data-testid="collapsedControl"] { display: none !important; }
      header { display: none !important; }
      footer { display: none !important; }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- Helpers ----------------
def norm_colname(c) -> str:
    if c is None:
        return ""
    s = str(c).replace("\n", " ").replace("\r", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("*", "")
    return s

def resolve_sheet_name(xls: pd.ExcelFile, preferred: str, fallbacks=()):
    if preferred in xls.sheet_names:
        return preferred
    pref_cf = preferred.casefold()
    for s in xls.sheet_names:
        if s.casefold() == pref_cf:
            return s
    # contains match
    for fb in [preferred] + list(fallbacks):
        fb_cf = fb.casefold()
        for s in xls.sheet_names:
            if fb_cf in s.casefold():
                return s
    raise ValueError(f"Could not find sheet '{preferred}'. Found sheets: {xls.sheet_names}")

def detect_mapping_columns(mapping: pd.DataFrame):
    cols = [norm_colname(c) for c in mapping.columns]
    def sig(c):
        return re.sub(r"[^a-z0-9]+", "", c.casefold())
    sigs = {c: sig(c) for c in cols}

    uz_candidates = [c for c in cols if ("uzio" in sigs[c]) and ("col" in sigs[c] or "column" in sigs[c] or "columns" in sigs[c])]
    adp_candidates = [c for c in cols if ("adp" in sigs[c]) and ("col" in sigs[c] or "column" in sigs[c] or "columns" in sigs[c])]

    if uz_candidates and adp_candidates:
        return uz_candidates[0], adp_candidates[0]

    # fall back: first two columns
    if len(cols) >= 2:
        return cols[0], cols[1]

    raise ValueError("Mapping Sheet must have at least 2 columns (UZIO col, ADP col).")

def is_blank(x) -> bool:
    return (
        x is None
        or (isinstance(x, float) and np.isnan(x))
        or (isinstance(x, str) and x.strip().lower() in {"", "nan", "none", "null"})
    )

def norm_key(v) -> str:
    if is_blank(v):
        return ""
    s = str(v).strip().replace("\u00A0", " ")
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s

def norm_boolish(x) -> str:
    if is_blank(x):
        return ""
    if isinstance(x, bool):
        return "yes" if x else "no"
    s = str(x).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return BOOL_TOKEN_MAP.get(s, s)

def try_float(x):
    if is_blank(x):
        return None
    if isinstance(x, (int, float, np.integer, np.floating)) and not (isinstance(x, float) and np.isnan(x)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "")
    try:
        return float(s)
    except Exception:
        return None

def norm_for_compare(field: str, value, side: str):
    """
    side: 'uzio' or 'adp'
    """
    # normalize boolean-ish first
    b = norm_boolish(value)
    if b in {"yes", "no"}:
        return b

    # cents -> dollars rule for specific fields
    if field in CENTS_FIELDS:
        v = try_float(value)
        if v is None:
            return 0.0 if field in ZERO_IF_BLANK_FIELDS else ""
        if side == "uzio":
            v = v / 100.0
        return float(v)

    # numeric blanks -> 0 for selected fields
    if field in ZERO_IF_BLANK_FIELDS:
        v = try_float(value)
        return 0.0 if v is None else float(v)

    # default: normalize as lowercase text with collapsed spaces
    if is_blank(value):
        return ""
    s = str(value).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def equal_norm(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(float(a), float(b), rel_tol=0, abs_tol=1e-9)
    return a == b

# ---------------- Core comparison ----------------
def run_active_only_comparison(file_bytes: bytes) -> tuple[bytes, dict]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio_sheet = resolve_sheet_name(xls, UZIO_SHEET_PREFERRED, ["Uzio Data", "UZIO"])
    adp_sheet  = resolve_sheet_name(xls, ADP_SHEET_PREFERRED,  ["Adp Data", "ADP"])
    map_sheet  = resolve_sheet_name(xls, MAP_SHEET_PREFERRED,  ["Mapping", "MAP"])

    uzio = pd.read_excel(xls, sheet_name=uzio_sheet, dtype=object)
    adp  = pd.read_excel(xls, sheet_name=adp_sheet, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=map_sheet, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    adp.columns  = [norm_colname(c) for c in adp.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    # Detect mapping columns
    uz_map_col, adp_map_col = detect_mapping_columns(mapping)

    mapping_valid = mapping.dropna(subset=[uz_map_col, adp_map_col]).copy()
    mapping_valid[uz_map_col] = mapping_valid[uz_map_col].astype(str).map(norm_colname)
    mapping_valid[adp_map_col] = mapping_valid[adp_map_col].astype(str).map(norm_colname)
    mapping_valid = mapping_valid[(mapping_valid[uz_map_col] != "") & (mapping_valid[adp_map_col] != "")]
    mapping_valid = mapping_valid.drop_duplicates(subset=[uz_map_col], keep="first")

    # Validate required columns
    if UZIO_KEY_PREFERRED not in uzio.columns:
        raise ValueError(f"UZIO key column '{UZIO_KEY_PREFERRED}' not found in UZIO sheet.")
    if "employment_status" not in uzio.columns:
        raise ValueError("UZIO sheet must contain 'employment_status' to filter ACTIVE employees.")
    if ADP_KEY_PREFERRED not in adp.columns:
        raise ValueError(f"ADP key column '{ADP_KEY_PREFERRED}' not found in ADP sheet.")

    # Active-only UZIO
    uzio_active = uzio[uzio["employment_status"].astype(str).str.upper().eq("ACTIVE")].copy()

    # Normalize keys
    uzio_active[UZIO_KEY_PREFERRED] = uzio_active[UZIO_KEY_PREFERRED].map(norm_key)
    adp[ADP_KEY_PREFERRED] = adp[ADP_KEY_PREFERRED].map(norm_key)

    uzio_active = uzio_active[uzio_active[UZIO_KEY_PREFERRED] != ""].drop_duplicates(subset=[UZIO_KEY_PREFERRED], keep="first")
    adp = adp[adp[ADP_KEY_PREFERRED] != ""].drop_duplicates(subset=[ADP_KEY_PREFERRED], keep="first")

    uz_idx = uzio_active.set_index(UZIO_KEY_PREFERRED, drop=False)
    adp_idx = adp.set_index(ADP_KEY_PREFERRED, drop=False)

    # Compare ONLY mapped fields (ignore everything else)
    uz_to_adp = dict(zip(mapping_valid[uz_map_col], mapping_valid[adp_map_col]))
    mapped_fields = [f for f in mapping_valid[uz_map_col].tolist() if f != UZIO_KEY_PREFERRED]

    rows = []
    missing_in_adp = []

    for emp_id in uzio_active[UZIO_KEY_PREFERRED].tolist():
        if emp_id not in adp_idx.index:
            missing_in_adp.append(emp_id)
            continue

        for field in mapped_fields:
            adp_col = uz_to_adp.get(field, "")
            if adp_col == "":
                continue

            uz_val = uz_idx.at[emp_id, field] if field in uz_idx.columns else ""
            adp_val = adp_idx.at[emp_id, adp_col] if adp_col in adp_idx.columns else ""

            uz_n = norm_for_compare(field, uz_val, side="uzio")
            adp_n = norm_for_compare(field, adp_val, side="adp")

            # Special equivalence mapping
            if (adp_n, uz_n) in SPECIAL_ADP_UZIO_EQUIVALENCE:
                status = "OK"
            else:
                if uz_n == "" and adp_n == "":
                    status = "OK"
                elif uz_n == "" and adp_n != "":
                    status = "UZIO_MISSING_VALUE"
                elif uz_n != "" and adp_n == "":
                    status = "ADP_MISSING_VALUE"
                else:
                    status = "OK" if equal_norm(uz_n, adp_n) else "MISMATCH"

            if status != "OK":
                rows.append({
                    "employee_id": emp_id,
                    "field": field,
                    "adp_column": adp_col,
                    "uzio_value_raw": uz_val,
                    "adp_value_raw": adp_val,
                    "uzio_value_normalized": uz_n,
                    "adp_value_normalized": adp_n,
                    "status": status,
                })

    mismatches_df = pd.DataFrame(rows)
    if len(mismatches_df):
        mismatches_df = mismatches_df.sort_values(["field", "employee_id"]).reset_index(drop=True)
        summary_by_field = (
            mismatches_df.groupby(["field", "status"])
            .size()
            .reset_index(name="count")
            .sort_values(["field", "status"])
            .reset_index(drop=True)
        )
    else:
        mismatches_df = pd.DataFrame(columns=[
            "employee_id","field","adp_column","uzio_value_raw","adp_value_raw",
            "uzio_value_normalized","adp_value_normalized","status"
        ])
        summary_by_field = pd.DataFrame(columns=["field","status","count"])

    # Build Excel output
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_by_field.to_excel(writer, sheet_name="Summary_By_Field", index=False)
        mismatches_df.to_excel(writer, sheet_name="Mismatches_Detail", index=False)

    metrics = {
        "uzio_sheet": uzio_sheet,
        "adp_sheet": adp_sheet,
        "mapping_sheet": map_sheet,
        "active_uzio_count": int(len(uzio_active)),
        "missing_in_adp_count": int(len(missing_in_adp)),
        "mismatch_rows": int(len(mismatches_df)),
    }
    return out.getvalue(), metrics

# ---------------- Streamlit UI ----------------
st.title(APP_TITLE)
st.write("Upload the Excel workbook (.xlsx) containing **UZIO Data**, **ADP Data**, and **Mapping Sheet**.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
run_btn = st.button("Run (Active Only)", type="primary", disabled=(uploaded_file is None))

if run_btn:
    try:
        with st.spinner("Comparing active UZIO employees against ADP..."):
            report_bytes, metrics = run_active_only_comparison(uploaded_file.getvalue())

        st.success("Report generated.")

        st.markdown(
            f"""
            **Sheets detected**
            - UZIO: `{metrics['uzio_sheet']}`
            - ADP: `{metrics['adp_sheet']}`
            - Mapping: `{metrics['mapping_sheet']}`

            **Counts**
            - Active employees in UZIO: **{metrics['active_uzio_count']}**
            - Missing in ADP: **{metrics['missing_in_adp_count']}**
            - Mismatch/Not-OK rows: **{metrics['mismatch_rows']}**
            """
        )

        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
