import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# UZIO vs ADP Comparison Tool (Streamlit)
# - Upload ONE Excel workbook with 3 tabs:
#     1) UZIO data
#     2) ADP data
#     3) Mapping sheet (UZIO col -> ADP col)
# - Generates an Excel report and provides download button
# - No table previews; no sidebar
#
# OUTPUT TABS:
#   - Summary
#   - Field_Summary_By_Status
#   - Mapping_ADP_Col_Missing
#   - Comparison_Detail_AllFields
#   - Mismatches_Only
#
# ADP is source of truth.
# =========================================================

APP_TITLE = "UZIO vs ADP Federal and State Withholing Comparison Audit"
OUTPUT_FILENAME = "UZIO_vs_ADP_Comparison_Report_ADP_SourceOfTruth.xlsx"

# Preferred sheet names (tool will also match case-insensitively and by "contains")
UZIO_SHEET_PREFERRED = "Uzio Data"
ADP_SHEET_PREFERRED = "ADP Data"
MAP_SHEET_PREFERRED = "Mapping Sheet"

# ---------- UI: Hide sidebar + Streamlit chrome ----------
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

# ---------- Helpers ----------
def norm_colname(c: str) -> str:
    if c is None:
        return ""
    c = str(c).replace("\n", " ").replace("\r", " ").replace("\u00A0", " ")
    c = re.sub(r"\s+", " ", c).strip()
    c = c.replace("*", "")
    c = c.strip('"').strip("'")
    return c

def norm_field_for_match(name: str) -> str:
    """Like norm_colname, but also turns underscores/hyphens into spaces for keyword checks."""
    s = norm_colname(name).casefold()
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def norm_blank(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    if isinstance(x, str) and x.strip().lower() in {"", "nan", "none", "null"}:
        return ""
    return x

def try_parse_date(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x).date().isoformat()
    if isinstance(x, str):
        s = x.strip()
        try:
            return pd.to_datetime(s, errors="raise").date().isoformat()
        except Exception:
            return s
    return str(x)

def digits_only(x):
    x = norm_blank(x)
    if x == "":
        return ""
    return re.sub(r"\D", "", str(x))

def norm_zip_first5(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)) and float(x).is_integer():
        s = str(int(x))
    else:
        s = re.sub(r"[^\d]", "", str(x).strip())
    if s == "":
        return ""
    if 0 < len(s) < 5:
        s = s.zfill(5)
    return s[:5]

NUMERIC_KEYWORDS = {"salary", "rate", "hours", "amount"}
DATE_KEYWORDS = {"date", "dob", "birth"}
SSN_KEYWORDS = {"ssn", "tax id"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
GENDER_KEYWORDS = {"gender"}

# ---------- Business constraints / normalization rules ----------
# ADP textual value -> UZIO coded value equivalences
SPECIAL_ADP_UZIO_EQUIVALENCE = {
    ("head of household", "federal_head_of_household"),
    ("single or married filing separately", "federal_single_or_married"),
    ("married filing jointly or qualifying surviving spouse", "federal_married_jointly"),
    ("single", "md_single"),
    ("married", "md_married"),
}

# UZIO stores these fields in cents; ADP stores in dollars
CENTS_FIELDS = {
    "fit_addl_withholding_per_pay_period",
    "fit_child_and_dependent_tax_credit",
    "fit_deductions_over_standard",
    "fit_other_income",
    "sit_addl_withholding_per_pay_period",
}

# Boolean normalization (UZIO: TRUE/FALSE, ADP: YES/NO)
BOOL_TOKEN_MAP = {
    "true": "yes",
    "false": "no",
    "yes": "yes",
    "no": "no",
    "y": "yes",
    "n": "no",
}

def norm_gender(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    if "female" in s or "woman" in s:
        return "female"
    if "male" in s or "man" in s:
        return "male"
    return s

def norm_value(x, field_name: str, side: str = ""):
    """Normalize values for comparison.

    side:
      - "uzio": apply UZIO-specific transforms (e.g., cents -> dollars)
      - "adp":  apply ADP-specific transforms
    """
    f = norm_field_for_match(field_name)
    x = norm_blank(x)
    if x == "":
        return ""

    # Boolean tokens: normalize TRUE/FALSE and YES/NO to 'yes'/'no'
    if isinstance(x, bool):
        return "yes" if x else "no"
    if isinstance(x, (int, np.integer)) and str(x) in ("0", "1"):
        return "yes" if int(x) == 1 else "no"
    if isinstance(x, (float, np.floating)) and float(x) in (0.0, 1.0):
        return "yes" if int(x) == 1 else "no"
    if isinstance(x, str):
        t = re.sub(r"\s+", " ", x.replace("\u00A0", " ")).strip().casefold()
        if t in BOOL_TOKEN_MAP:
            return BOOL_TOKEN_MAP[t]

    # UZIO cents -> dollars for specific fields
    field_cf = norm_colname(field_name).casefold()
    if field_cf in CENTS_FIELDS:
        # Parse numeric on both sides; only UZIO divides by 100
        if isinstance(x, (int, float, np.integer, np.floating)):
            v = float(x)
        else:
            s = str(x).strip().replace(",", "").replace("$", "")
            try:
                v = float(s)
            except Exception:
                return re.sub(r"\s+", " ", str(x).strip()).casefold()
        if side.casefold() == "uzio":
            v = v / 100.0
        return float(v)

    if any(k in f for k in GENDER_KEYWORDS):
        return norm_gender(x)

    if any(k in f for k in SSN_KEYWORDS):
        return digits_only(x)

    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    if any(k in f for k in NUMERIC_KEYWORDS):
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        if isinstance(x, str):
            s = x.strip().replace(",", "").replace("$", "")
            try:
                return float(s)
            except Exception:
                return re.sub(r"\s+", " ", x.strip()).casefold()

    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()

    return str(x).casefold()

def norm_emp_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

# ---------- Rule helpers ----------
def is_termination_reason_field(field_name: str) -> bool:
    return "termination reason" in norm_field_for_match(field_name)

def is_employment_status_field(field_name: str) -> bool:
    s = norm_field_for_match(field_name)
    return ("employment status" in s) or (s == "employment_status")

def status_contains_any(s: str, needles) -> bool:
    s = ("" if s is None else str(s)).casefold()
    return any(n in s for n in needles)

def uzio_is_active(uz_norm: str) -> bool:
    s = ("" if uz_norm is None else str(uz_norm)).casefold()
    return s == "active" or s.startswith("active")

def uzio_is_terminated(uz_norm: str) -> bool:
    s = ("" if uz_norm is None else str(uz_norm)).casefold()
    return s == "terminated" or s.startswith("terminated")

ALLOWED_TERM_REASONS = {
    "quit without notice",
    "no reason given",
    "misconduct",
    "abandoned job",
    "advancement (better job with higher pay)",
    "no-show (never started employment)",
    "performance",
    "personal",
    "scheduling conflicts (schedules don't work)",
    "attendance",
}

def normalize_reason_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = s.replace("’", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip('"').strip("'")
    return s.casefold()

def normalize_paytype_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def paytype_bucket(paytype_norm: str) -> str:
    s = ("" if paytype_norm is None else str(paytype_norm)).casefold()
    if "hour" in s:
        return "hourly"
    if "salary" in s or "salaried" in s:
        return "salaried"
    return ""

def is_annual_salary_field(field_name: str) -> bool:
    return "annual salary" in norm_field_for_match(field_name)

def is_hourly_rate_field(field_name: str) -> bool:
    f = norm_field_for_match(field_name)
    return ("hourly pay rate" in f) or ("hourly rate" in f)

# ---------- Guardrail: prevent ACTIVE/TERMINATED/RETIRED values leaking into non-status fields ----------
EMP_STATUS_TOKENS = {"active", "terminated", "retired"}

def field_allows_emp_status_value(field_name: str) -> bool:
    f = norm_field_for_match(field_name)
    return (f == "status") or ("employment status" in f)

def cleanse_uzio_value_for_field(field_name: str, uz_val):
    if norm_blank(uz_val) == "":
        return uz_val
    s = str(uz_val).strip().casefold()
    if (s in EMP_STATUS_TOKENS) and (not field_allows_emp_status_value(field_name)):
        return ""
    return uz_val

# ---------- Pay Type equivalence (UZIO Salaried == ADP Salary) ----------
def is_pay_type_field(field_name: str) -> bool:
    f = norm_field_for_match(field_name)
    return f == "pay type" or ("pay type" in f)

def normalize_paytype_for_compare(x) -> str:
    s = normalize_paytype_text(x)
    if s in {"salary", "salaried"}:
        return "salaried"
    if s in {"hourly", "hour"}:
        return "hourly"
    return s

# ---------- Sheet + Mapping detection ----------
def resolve_sheet_name(xls: pd.ExcelFile, preferred: str, fallbacks):
    if preferred in xls.sheet_names:
        return preferred
    pref_cf = preferred.casefold()
    for s in xls.sheet_names:
        if s.casefold() == pref_cf:
            return s
    for fb in [preferred] + list(fallbacks):
        if fb in xls.sheet_names:
            return fb
        fb_cf = fb.casefold()
        for s in xls.sheet_names:
            if s.casefold() == fb_cf:
                return s
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
    if len(cols) >= 2:
        return cols[0], cols[1]
    raise ValueError("Mapping sheet must have at least 2 columns: UZIO column name and ADP column name.")

def detect_key_mapping(mapping_valid: pd.DataFrame, uzio_cols, adp_cols, uz_map_col, adp_map_col):
    def score(uz, adp):
        uz_s = norm_field_for_match(uz)
        adp_s = norm_field_for_match(adp)
        sc = 0
        if "employee" in uz_s: sc += 4
        if uz_s in {"employee id", "employee_id"}: sc += 6
        if "id" in uz_s: sc += 2
        if "associate id" in adp_s: sc += 6
        if adp_s.endswith("id") or " id" in adp_s: sc += 2
        if uz in uzio_cols: sc += 2
        if adp in adp_cols: sc += 2
        return sc

    best = None
    best_sc = -1
    for _, r in mapping_valid.iterrows():
        uz = r[uz_map_col]
        adp = r[adp_map_col]
        if uz not in uzio_cols or adp not in adp_cols:
            continue
        sc = score(uz, adp)
        if sc > best_sc:
            best_sc = sc
            best = (uz, adp)
    if best is None:
        raise ValueError("Could not detect key mapping. Ensure the mapping includes an ID field mapped between UZIO and ADP.")
    return best

# ---------- Core compare ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio_sheet = resolve_sheet_name(xls, UZIO_SHEET_PREFERRED, ["UZIO Data", "UZIO_DATA", "uzio data", "UZIO"])
    adp_sheet  = resolve_sheet_name(xls, ADP_SHEET_PREFERRED,  ["Adp Data", "ADP_DATA", "adp data", "ADP"])
    map_sheet  = resolve_sheet_name(xls, MAP_SHEET_PREFERRED,  ["mapping sheet", "Mapping", "MAP"])

    uzio = pd.read_excel(xls, sheet_name=uzio_sheet, dtype=object)
    adp = pd.read_excel(xls, sheet_name=adp_sheet, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=map_sheet, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    adp.columns = [norm_colname(c) for c in adp.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    uz_map_col, adp_map_col = detect_mapping_columns(mapping)

    mapping[uz_map_col] = mapping[uz_map_col].map(norm_colname)
    mapping[adp_map_col] = mapping[adp_map_col].map(norm_colname)

    mapping_valid = mapping.dropna(subset=[uz_map_col, adp_map_col]).copy()
    mapping_valid = mapping_valid[(mapping_valid[uz_map_col] != "") & (mapping_valid[adp_map_col] != "")]
    mapping_valid = mapping_valid.drop_duplicates(subset=[uz_map_col], keep="first").copy()

    UZIO_KEY, ADP_KEY = detect_key_mapping(mapping_valid, set(uzio.columns), set(adp.columns), uz_map_col, adp_map_col)

    if UZIO_KEY not in uzio.columns:
        raise ValueError(f"UZIO key column '{UZIO_KEY}' not found in UZIO sheet.")
    if ADP_KEY not in adp.columns:
        raise ValueError(f"ADP key column '{ADP_KEY}' not found in ADP sheet.")

    uzio[UZIO_KEY] = norm_emp_key_series(uzio[UZIO_KEY])
    adp[ADP_KEY] = norm_emp_key_series(adp[ADP_KEY])

    uzio = uzio.drop_duplicates(subset=[UZIO_KEY], keep="first").copy()
    adp = adp.drop_duplicates(subset=[ADP_KEY], keep="first").copy()

    uzio_keys = set(uzio[UZIO_KEY].dropna().astype(str).str.strip()) - {""}
    adp_keys = set(adp[ADP_KEY].dropna().astype(str).str.strip()) - {""}
    all_keys = sorted(uzio_keys.union(adp_keys))

    uzio_idx = uzio.set_index(UZIO_KEY, drop=False)
    adp_idx = adp.set_index(ADP_KEY, drop=False)

    uz_to_adp = dict(zip(mapping_valid[uz_map_col], mapping_valid[adp_map_col]))
    mapped_fields = [f for f in mapping_valid[uz_map_col].tolist() if f != UZIO_KEY]

    mapping_missing_adp_col = mapping_valid[~mapping_valid[adp_map_col].isin(adp.columns)].copy()

    # UZIO Employment Status column for context column
    uzio_employment_status_col = None
    for c in uzio.columns:
        if norm_field_for_match(c) == "employment status":
            uzio_employment_status_col = c
            break
    if uzio_employment_status_col is None:
        for c in uzio.columns:
            if "employment status" in norm_field_for_match(c):
                uzio_employment_status_col = c
                break

    def get_uzio_employment_status(emp_id: str) -> str:
        if uzio_employment_status_col is None:
            return ""
        if emp_id in uzio_idx.index and uzio_employment_status_col in uzio_idx.columns:
            v = uzio_idx.at[emp_id, uzio_employment_status_col]
            return "" if norm_blank(v) == "" else str(v)
        return ""

    # Pay Type mapping (prefer ADP) for context column and pay-type exceptions
    paytype_rows = mapping_valid[mapping_valid[uz_map_col].map(lambda x: "pay type" in norm_field_for_match(x))]
    UZIO_PAYTYPE_COL = paytype_rows.iloc[0][uz_map_col] if len(paytype_rows) else None
    ADP_PAYTYPE_COL  = paytype_rows.iloc[0][adp_map_col] if len(paytype_rows) else None

    def get_employee_pay_type(emp_id: str, adp_exists: bool, uz_exists: bool) -> str:
        if ADP_PAYTYPE_COL and adp_exists and (ADP_PAYTYPE_COL in adp_idx.columns):
            v = adp_idx.at[emp_id, ADP_PAYTYPE_COL]
            if norm_blank(v) != "":
                return str(v)
        if UZIO_PAYTYPE_COL and uz_exists and (UZIO_PAYTYPE_COL in uzio_idx.columns):
            v = uzio_idx.at[emp_id, UZIO_PAYTYPE_COL]
            if norm_blank(v) != "":
                return str(v)
        return ""

    rows = []
    for emp_id in all_keys:
        uz_exists = emp_id in uzio_idx.index
        adp_exists = emp_id in adp_idx.index

        uz_emp_status = get_uzio_employment_status(emp_id)
        emp_paytype = get_employee_pay_type(emp_id, adp_exists=adp_exists, uz_exists=uz_exists)
        emp_pay_bucket = paytype_bucket(normalize_paytype_text(emp_paytype))

        for field in mapped_fields:
            adp_col = uz_to_adp.get(field, "")

            uz_val_raw = uzio_idx.at[emp_id, field] if (uz_exists and field in uzio_idx.columns) else ""
            uz_val = cleanse_uzio_value_for_field(field, uz_val_raw)

            adp_val = adp_idx.at[emp_id, adp_col] if (adp_exists and (adp_col in adp_idx.columns)) else ""

            # Employee missing logic (ADP truth)
            if not adp_exists and uz_exists:
                status = "MISSING_IN_ADP"
            elif adp_exists and not uz_exists:
                status = "MISSING_IN_UZIO"
            elif adp_exists and uz_exists and (adp_col not in adp.columns):
                status = "ADP_COLUMN_MISSING"
            else:
                # Pay Type equivalence
                if is_pay_type_field(field):
                    uz_pt = normalize_paytype_for_compare(uz_val)
                    adp_pt = normalize_paytype_for_compare(adp_val)

                    if (uz_pt == adp_pt) or (uz_pt == "" and adp_pt == ""):
                        status = "OK"
                    elif uz_pt == "" and adp_pt != "":
                        status = "UZIO_MISSING_VALUE"
                    elif uz_pt != "" and adp_pt == "":
                        status = "ADP_MISSING_VALUE"
                    else:
                        status = "MISMATCH"
                else:
                    uz_n = norm_value(uz_val, field, side="uzio")
                    adp_n = norm_value(adp_val, field, side="adp")

                    # Business constraint: ADP value text maps to UZIO coded value (treat as match)
                    if (adp_n, uz_n) in SPECIAL_ADP_UZIO_EQUIVALENCE:
                        status = "OK"
                    # Employment Status special rule
                    elif is_employment_status_field(field) and adp_n != "":
                        adp_is_term_or_ret = status_contains_any(adp_n, ["terminated", "retired"])
                        if adp_is_term_or_ret:
                            if uz_n == "":
                                status = "UZIO_MISSING_VALUE"
                            elif uzio_is_active(uz_n):
                                status = "MISMATCH"
                            elif uzio_is_terminated(uz_n):
                                status = "OK"
                            else:
                                status = "MISMATCH"
                        else:
                            if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                                status = "OK"
                            elif uz_n == "" and adp_n != "":
                                status = "UZIO_MISSING_VALUE"
                            elif uz_n != "" and adp_n == "":
                                status = "ADP_MISSING_VALUE"
                            else:
                                status = "MISMATCH"

                    # Termination Reason special rule
                    elif is_termination_reason_field(field):
                        uz_reason = normalize_reason_text(uz_val)
                        adp_reason = normalize_reason_text(adp_val)

                        if uz_reason == "other" and adp_reason in ALLOWED_TERM_REASONS:
                            status = "OK"
                        else:
                            if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                                status = "OK"
                            elif uz_n == "" and adp_n != "":
                                status = "UZIO_MISSING_VALUE"
                            elif uz_n != "" and adp_n == "":
                                status = "ADP_MISSING_VALUE"
                            else:
                                status = "MISMATCH"

                    # Default field compare
                    else:
                        if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                            status = "OK"
                        elif uz_n == "" and adp_n != "":
                            status = "UZIO_MISSING_VALUE"
                        elif uz_n != "" and adp_n == "":
                            status = "ADP_MISSING_VALUE"
                        else:
                            status = "MISMATCH"

                        # PayType exceptions overriding UZIO_MISSING_VALUE for amount/rate fields
                        if status == "UZIO_MISSING_VALUE":
                            if emp_pay_bucket == "hourly" and is_annual_salary_field(field):
                                status = "OK"
                            elif emp_pay_bucket == "salaried" and is_hourly_rate_field(field):
                                status = "OK"

            rows.append({
                "Employee ID": emp_id,
                "UZIO Employment Status": uz_emp_status,
                "Pay Type": emp_paytype,
                "Field": field,
                "UZIO_Value": uz_val,
                "ADP_Value": adp_val,
                "ADP_SourceOfTruth_Status": status
            })

    comparison_detail = pd.DataFrame(rows)
    mismatches_only = comparison_detail[comparison_detail["ADP_SourceOfTruth_Status"] != "OK"].copy()

    # Field Summary By Status
    cols_needed = [
        "OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "ADP_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_ADP",
        "ADP_COLUMN_MISSING",
    ]

    pivot = comparison_detail.pivot_table(
        index="Field",
        columns="ADP_SourceOfTruth_Status",
        values="Employee ID",
        aggfunc="count",
        fill_value=0
    )

    for c in cols_needed:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["Total"] = pivot.sum(axis=1)
    pivot["OK"] = pivot["OK"].astype(int)
    pivot["NOT_OK"] = (pivot["Total"] - pivot["OK"]).astype(int)

    field_summary_by_status = pivot.reset_index()[[
        "Field",
        "Total",
        "OK",
        "NOT_OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "ADP_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_ADP",
        "ADP_COLUMN_MISSING",
    ]]

    # Summary metrics
    summary = pd.DataFrame({
        "Metric": [
            "UZIO sheet name",
            "ADP sheet name",
            "Mapping sheet name",
            "Key mapping (UZIO -> ADP)",
            "Employees in UZIO sheet",
            "Employees in ADP sheet",
            "Employees present in both",
            "Employees missing in ADP (UZIO only)",
            "Employees missing in UZIO (ADP only)",
            "Mapped fields total (from mapping sheet)",
            "Mapped fields with ADP column missing",
            "Total comparison rows (employees x mapped fields)",
            "Total NOT OK rows"
        ],
        "Value": [
            uzio_sheet,
            adp_sheet,
            map_sheet,
            f"{UZIO_KEY} -> {ADP_KEY}",
            len(uzio_keys),
            len(adp_keys),
            len(uzio_keys.intersection(adp_keys)),
            len(uzio_keys - adp_keys),
            len(adp_keys - uzio_keys),
            len(mapped_fields),
            mapping_missing_adp_col.shape[0],
            comparison_detail.shape[0],
            mismatches_only.shape[0]
        ]
    })

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Field_Summary_By_Status", index=False)
        mapping_missing_adp_col.to_excel(writer, sheet_name="Mapping_ADP_Col_Missing", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison_Detail_AllFields", index=False)
        mismatches_only.to_excel(writer, sheet_name="Mismatches_Only", index=False)

    return out.getvalue()

# ---------- Minimal UI ----------
st.title(APP_TITLE)
st.write("Upload the Excel workbook (.xlsx) that contains: UZIO Data, ADP Data, and Mapping Sheet.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
run_btn = st.button("Run Audit", type="primary", disabled=(uploaded_file is None))

if run_btn:
    try:
        with st.spinner("Running audit..."):
            report_bytes = run_comparison(uploaded_file.getvalue())

        st.success("Report generated.")
        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
