import io
import re
from datetime import datetime
from typing import Dict, Any, Tuple, Optional, List

import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# 1) UZIO enum -> ADP filing status label mapping (your mapping)
#    Structure: {PREFIX: {SUFFIX: ADP_LABEL}}
#    Example UZIO enum: "IL_SINGLE" => prefix="IL", suffix="SINGLE"
# ============================================================
FILING_STATUS_MAP: Dict[str, Dict[str, str]] = {
    "FEDERAL": {
        "SINGLE": "Single",
        "MARRIED": "Married",
        "MARRIED_SINGLE": "Married but withhold as Single",
        "SINGLE_OR_MARRIED": "Single or Married filing separately",
        "MARRIED_JOINTLY": "Married filing jointly or Qualifying surviving spouse",
        "HEAD_OF_HOUSEHOLD": "Head of household",
    },
    "MD": {"SINGLE": "Single", "MARRIED": "Married", "MARRIED_SINGLE": "Married but withhold at single rate"},
    "DC": {
        "SINGLE": "Single",
        "MARRIED_DP_JOINTLY": "Married/Domestic Partners Filing Jointly",
        "MARRIED_SEPARATELY": "Married Filing Separately",
        "HEAD_OF_HOUSEHOLD": "Head of Household",
        "MARRIED_DP_SEPARATELY": "Married/Domestic Partners Filing Separately",
    },
    "NM": {
        "SINGLE": "Single or Married filing separately",
        "MARRIED": "Married filing jointly or Qualifying Surviving Spouse",
        "MARRIED_SINGLE": "Married but withhold as Single",
        "HEAD_OF_HOUSEHOLD": "Head of Household",
    },
    "MS": {"SINGLE": "Single", "HEAD_OF_HOUSEHOLD": "Head of Family", "M1": "Married (Spouse NOT employed)", "M2": "Married (Spouse is employed)"},
    "MO": {"SINGLE": "Single or Married Spouse Works or Married Filing Separate", "MARRIED": "Married (Spouse does not work)", "HEAD_OF_HOUSEHOLD": "Head of Household"},
    "AL": {
        "NO_PERSONAL_EXEMPTION": "No Personal Exemption",
        "SINGLE": "Single",
        "MARRIED": "Married",
        "MARRIED_SEPARATELY": "Married Filing Separately",
        "HEAD_OF_HOUSEHOLD": "Head of Family",
    },
    "DE": {"MARRIED": "Married", "SINGLE": "Single", "MARRIED_SINGLE_RATE": "Married but Withhold as Single"},
    "OK": {"MARRIED": "Married", "SINGLE": "Single", "MARRIED_SINGLE_RATE": "Married but Withhold as Single", "NRA": "Non-Resident Alien"},
    "NC": {"HEAD_OF_HOUSEHOLD": "Head of Household", "MARRIED": "Married Filing Jointly or Surviving Spouse", "SINGLE": "Single or Married Filing Separately"},
    "SC": {"MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate", "MARRIED": "Married", "SINGLE": "Single"},
    "UT": {"SINGLE": "Single or Married filing separately", "MARRIED": "Married filing jointly or Qualifying widow(er)", "HEAD_OF_HOUSEHOLD": "Head of Household"},
    "GA": {
        "SINGLE": "Single",
        "SEPARATE_MARRIED_JOINT_BOTH_WORKING": "Married Filing Separate or Married Filing Joint both spouses working",
        "MARRIED_JOINT_ONE_WORKING": "Married Filing Joint one spouse working",
        "HEAD_OF_HOUSEHOLD": "Head of Household",
    },
    "WI": {"SINGLE": "Single", "MARRIED": "Married", "MARRIED_SINGLE_RATE": "Married but withhold at higher single rate"},
    "KS": {"SINGLE": "Single", "JOINT": "Joint"},
    "VT": {
        "SINGLE": "Single",
        "MARRIED": "Married/Civil Union Filing Jointly",
        "MARRIED_FILING_SEPERATELY": "Married/Civil Union Filing Separately",
        "MARRIED_SINGLE_RATE": "Married, but withhold at higher single rate",
    },
    "NJ": {
        "SINGLE": "Single",
        "MARRIED_DP_JOINTLY": "Married/Civil Union Couple Joint",
        "MARRIED_SEPARATELY": "Married/Civil Union Partner Separate",
        "HEAD_OF_HOUSEHOLD": "Head of Household",
        "QUALIFIED_WIDOW": "Qualifying Widow(er)/Surviving Civil Union Partner",
    },
    "CA": {"HEAD_OF_HOUSEHOLD": "Head of Household", "MARRIED": "Married (one income)", "SINGLE": "Single or Married (with two or more incomes)"},
    "MN": {
        "SINGLE": "Single, Married but legally separated or Spouse is a nonresident alien",
        "MARRIED": "Married",
        "MARRIED_SINGLE_RATE": "Married but withhold at higher single rate",
    },
    "IA": {"OTHER": "Other (Including Single)", "HEAD_OF_HOUSEHOLD": "Head of Household", "MARRIED_JOINTLY": "Married filing jointly", "QUALIFIED_SPOUSE": "Qualifying Surviving Spouse"},
    "ME": {"SINGLE": "Single or Head of Household", "MARRIED": "Married", "MARRIED_SINGLE_RATE": "Married but withhold at higher single rate", "NON_RESIDENT_ALIEN": "Nonresident alien"},
    "NY": {"MARRIED_WITHHOLD_SINGLE": "Married but withhold as Single", "SINGLE": "Single", "MARRIED": "Married", "HEAD_OF_HOUSEHOLD": "Head of Household"},
    "NE": {"SINGLE": "Single", "MARRIED": "Married Filing Jointly or Qualifying Widow(er)"},
    "LA": {"NO_DEDUCTION": "No Deduction", "SINGLE_OR_MARRIED": "Single or married filing separately", "MARRIED_FILING_JOINTLY_HOH": "Married filing jointly, qualifying surviving spouse, or head of household"},
    "OR": {"SINGLE": "Single", "MARRIED": "Married", "MARRIED_SINGLE_RATE": "Married but withhold at higher single rate"},
    "ND": {
        "SINGLE": "Single",
        "MARRIED": "Married",
        "MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
        "SINGLE_MARRIED_SEPARATELY": "Single or Married filing separately",
        "HEAD_OF_HOUSEHOLD": "Head of household",
        "MARRIED_JOINTLY": "Married filing jointly  or Qualifying Surviving Spouse",
    },
    "ID": {"SINGLE": "Single", "MARRIED": "Married", "MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate"},
    "CO": {
        "SINGLE_OR_MARRIED_SEPARATELY": "Single or Married filing separately",
        "MARRIED_JOINTLY": "Married filing jointly",
        "HEAD_OF_HOUSEHOLD": "Head of household",
        "SINGLE": "Single",
        "MARRIED": "Married",
        "MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
    },
    "HI": {"SINGLE": "Single", "MARRIED": "Married", "MARRIED_SINGLE_RATE": "Married but Withhold at higher single rate", "DISABLED": "Certified disabled person", "NMS": "Nonresident Military Spouse"},
    "MT": {"SINGLE": "Single or Married filing separately", "MARRIED": "Married filing jointly or qualifying surviving spouse", "HEAD_OF_HOUSEHOLD": "Head of household"},
    "AR": {"SINGLE": "Single", "MARRIED_FILING_JOINTLY": "Married Filing Jointly", "HOH": "Head of Household"},
}

# ============================================================
# 2) Money fields where UZIO is in cents (÷100) but ADP is dollars
# ============================================================
DEFAULT_CENTS_FIELDS = {
    "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
    "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",
    "FIT_DEDUCTIONS_OVER_STANDARD",
    "FIT_OTHER_INCOME",
    "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
}

# ============================================================
# 3) Special rule: SIT allowances
#    If SIT_TOTAL_ALLOWANCES missing, compare BASIC + ADDITIONAL
# ============================================================
SIT_TOTAL_ALLOWANCES_FIELD = "SIT_TOTAL_ALLOWANCES"

SIT_TOTAL_ALLOWANCES_CANDS = ["SIT_TOTAL_ALLOWANCES", "SIT_ALLOWANCES", "STATE_ALLOWANCES"]

SIT_BASIC_ALLOWANCE_CANDS = [
    "SIT_BASIC_ALLOWANCE", "SIT_BASIC_ALLOWANCES", "STATE_BASIC_ALLOWANCE", "STATE_BASIC_ALLOWANCES",
    "SIT_ALLOWANCE_BASIC", "SIT_ALLOWANCES_BASIC"
]
SIT_ADDITIONAL_ALLOWANCES_CANDS = [
    "SIT_ADDITIONAL_ALLOWANCES", "SIT_ADDITIONAL_ALLOWANCE",
    "STATE_ADDITIONAL_ALLOWANCES", "STATE_ADDITIONAL_ALLOWANCE",
    "SIT_ALLOWANCE_ADDITIONAL", "SIT_ALLOWANCES_ADDITIONAL"
]

# ============================================================
# 4) Column name detection candidates
# ============================================================
KEY_CANDIDATES = ["EMPLOYEE_ID", "ASSOCIATE_ID", "EE_ID", "EMP_ID", "WORKER_ID", "EMPLOYEEID"]
STATUS_CANDIDATES = ["EMPLOYMENT_STATUS", "STATUS", "EE_STATUS"]

NAME_FIRST_CANDIDATES = ["FIRST_NAME", "EMPLOYEE_FIRST_NAME", "EE_FIRST_NAME", "WORKER_FIRST_NAME"]
NAME_LAST_CANDIDATES  = ["LAST_NAME", "EMPLOYEE_LAST_NAME", "EE_LAST_NAME", "WORKER_LAST_NAME"]

# ============================================================
# 5) Canonical field specs (tries to map ADP columns to UZIO columns)
#    Everything else (optional) can be compared by relevant common columns.
# ============================================================
FIELD_SPECS = [
    ("FIT_FILING_STATUS", "filing_status",
     ["FIT_FILING_STATUS", "FEDERAL_FILING_STATUS", "FILING_STATUS_FEDERAL", "FIT_STATUS"],
     ["FIT_FILING_STATUS", "FEDERAL_FILING_STATUS", "FEDERAL_FILINGSTATUS", "FIT_STATUS"]),
    ("SIT_FILING_STATUS", "filing_status",
     ["SIT_FILING_STATUS", "STATE_FILING_STATUS", "FILING_STATUS_STATE", "SIT_STATUS"],
     ["SIT_FILING_STATUS", "STATE_FILING_STATUS", "SIT_STATUS", "SIT_FILINGSTATUS"]),

    ("FIT_EXEMPT", "boolean",
     ["FIT_EXEMPT", "FIT_WITHHOLDING_EXEMPT", "DO_NOT_CALCULATE_FIT", "FEDERAL_EXEMPT"],
     ["FIT_EXEMPT", "FIT_WITHHOLDING_EXEMPT", "DO_NOT_CALCULATE_FIT", "FEDERAL_EXEMPT"]),
    ("SIT_EXEMPT", "boolean",
     ["SIT_EXEMPT", "SIT_WITHHOLDING_EXEMPT", "DO_NOT_CALCULATE_SIT", "STATE_EXEMPT"],
     ["SIT_EXEMPT", "SIT_WITHHOLDING_EXEMPT", "DO_NOT_CALCULATE_SIT", "STATE_EXEMPT"]),

    ("FIT_MULTIPLE_JOBS", "boolean",
     ["FIT_MULTIPLE_JOBS", "MULTIPLE_JOBS", "TWO_JOBS", "HIGHER_WITHHOLDING"],
     ["FIT_MULTIPLE_JOBS", "MULTIPLE_JOBS", "TWO_JOBS", "HIGHER_WITHHOLDING"]),

    ("FIT_DEPENDENTS_AMOUNT", "money",
     ["FIT_DEPENDENTS_AMOUNT", "DEPENDENTS_AMOUNT", "FIT_DEPENDENTS"],
     ["FIT_DEPENDENTS_AMOUNT", "DEPENDENTS_AMOUNT", "FIT_DEPENDENTS"]),

    ("FIT_OTHER_INCOME", "money",
     ["FIT_OTHER_INCOME", "OTHER_INCOME"],
     ["FIT_OTHER_INCOME", "OTHER_INCOME"]),

    ("FIT_DEDUCTIONS_OVER_STANDARD", "money",
     ["FIT_DEDUCTIONS_OVER_STANDARD", "DEDUCTIONS_OVER_STANDARD"],
     ["FIT_DEDUCTIONS_OVER_STANDARD", "DEDUCTIONS_OVER_STANDARD"]),

    ("FIT_CHILD_AND_DEPENDENT_TAX_CREDIT", "money",
     ["FIT_CHILD_AND_DEPENDENT_TAX_CREDIT", "CHILD_DEP_CREDIT", "CHILDREN_CREDIT"],
     ["FIT_CHILD_AND_DEPENDENT_TAX_CREDIT", "CHILD_DEP_CREDIT", "CHILDREN_CREDIT"]),

    ("FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "money",
     ["FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "FIT_ADDITIONAL_WITHHOLDING", "ADDITIONAL_FIT"],
     ["FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "FIT_ADDITIONAL_WITHHOLDING", "ADDITIONAL_FIT"]),

    ("SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "money",
     ["SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "SIT_ADDITIONAL_WITHHOLDING", "ADDITIONAL_SIT"],
     ["SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "SIT_ADDITIONAL_WITHHOLDING", "ADDITIONAL_SIT"]),
]

# ============================================================
# 6) Normalization helpers
# ============================================================
YES = {"YES", "Y", "TRUE", "T", "1"}
NO  = {"NO", "N", "FALSE", "F", "0"}

def normalize_colname(c: str) -> str:
    return re.sub(r"\s+", "_", str(c).strip().upper())

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_colname(c) for c in df.columns]
    return df

def detect_long_format(df: pd.DataFrame) -> bool:
    cols = set(df.columns)
    return ("WITHHOLDING_FIELD_KEY" in cols and "WITHHOLDING_FIELD_VALUE" in cols)

def pivot_long_to_wide(df: pd.DataFrame, id_cols: List[str]) -> pd.DataFrame:
    wide = (
        df.pivot_table(
            index=id_cols,
            columns="WITHHOLDING_FIELD_KEY",
            values="WITHHOLDING_FIELD_VALUE",
            aggfunc=lambda x: next((v for v in x if str(v).strip() != ""), "")
        )
        .reset_index()
    )
    wide.columns = [normalize_colname(c) for c in wide.columns]
    return wide

def norm_text_for_compare(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    s = str(v).strip()
    if s == "":
        return ""
    s = s.lower()
    # treat punctuation (including /) as insignificant
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def to_bool_or_blank(v: Any) -> Any:
    s = str(v).strip().upper() if v is not None else ""
    if s == "":
        return ""
    if s in YES:
        return True
    if s in NO:
        return False
    return s  # unknown representation remains as-is

def to_int_allowances(v: Any) -> Any:
    """
    Allowances are numeric.
    If TRUE/FALSE/Yes/No present, map TRUE->1, FALSE->0.
    """
    s = str(v).strip().upper() if v is not None else ""
    if s == "":
        return ""
    if s in YES:
        return 1
    if s in NO:
        return 0
    try:
        n = float(str(v).replace(",", ""))
        return int(n) if n.is_integer() else n
    except Exception:
        return s

def to_int_or_zero(v: Any) -> int:
    """
    For allowance summing: blank -> 0, True/False -> 1/0, numeric -> int
    """
    if v is None:
        return 0
    s = str(v).strip()
    if s == "":
        return 0
    u = s.upper()
    if u in YES:
        return 1
    if u in NO:
        return 0
    try:
        n = float(s.replace(",", ""))
        return int(n) if n.is_integer() else int(round(n))
    except Exception:
        return 0

def to_money_str(v: Any, uzio_is_cents: bool) -> str:
    s = str(v).strip() if v is not None else ""
    if s == "":
        return ""
    # if boolean sneaks in, normalize to 1/0
    if s.upper() in YES:
        s = "1"
    if s.upper() in NO:
        s = "0"
    try:
        n = float(s.replace(",", ""))
        if uzio_is_cents:
            n = n / 100.0
        return f"{n:.2f}"
    except Exception:
        return s

def parse_uzio_enum_to_adp_label(enum_val: Any) -> str:
    """
    Convert UZIO enum like 'IL_SINGLE' or 'FEDERAL_HEAD_OF_HOUSEHOLD' to ADP label using mapping.
    If no mapping found, fallback to readable transformation.
    """
    raw = str(enum_val).strip() if enum_val is not None else ""
    if raw == "":
        return ""
    u = raw.strip().upper()
    if "_" not in u:
        return raw  # already a label-ish value
    prefix, suffix = u.split("_", 1)
    if prefix in FILING_STATUS_MAP and suffix in FILING_STATUS_MAP[prefix]:
        return FILING_STATUS_MAP[prefix][suffix]
    # fallback if the state prefix isn't in table
    return suffix.replace("_", " ").title()

def resolve_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def pick_first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    return None

def resolve_allowance_cols(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    total = resolve_column(df, SIT_TOTAL_ALLOWANCES_CANDS)
    basic = resolve_column(df, SIT_BASIC_ALLOWANCE_CANDS)
    addl  = resolve_column(df, SIT_ADDITIONAL_ALLOWANCES_CANDS)
    return total, basic, addl

def compute_sit_total_allowances(row: pd.Series, total_col: Optional[str], basic_col: Optional[str], addl_col: Optional[str]) -> Any:
    """
    Preference:
      1) Use SIT_TOTAL_ALLOWANCES if present AND non-blank
      2) Else BASIC + ADDITIONAL (missing parts treated as 0)
      3) If nothing exists at all, return blank
    """
    if total_col and total_col in row.index:
        raw_total = row.get(total_col, "")
        if str(raw_total).strip() != "":
            return to_int_allowances(raw_total)

    has_basic = bool(basic_col and basic_col in row.index)
    has_addl  = bool(addl_col and addl_col in row.index)

    if not has_basic and not has_addl:
        return ""

    basic_val = to_int_or_zero(row.get(basic_col, "")) if has_basic else 0
    addl_val  = to_int_or_zero(row.get(addl_col, "")) if has_addl else 0
    return basic_val + addl_val

def allowance_col_label(total_col: Optional[str], basic_col: Optional[str], addl_col: Optional[str]) -> str:
    if total_col:
        return total_col
    parts = [c for c in [basic_col, addl_col] if c]
    return " + ".join(parts) if parts else ""

def normalize_pair(field: str, ftype: str, adp_val: Any, uzio_val: Any, cents_fields: set) -> Tuple[Any, Any, str]:
    """
    Returns (adp_norm, uzio_norm, rule_used)
    """
    if ftype == "filing_status":
        adp_label = str(adp_val).strip() if adp_val is not None else ""
        uzio_label = parse_uzio_enum_to_adp_label(uzio_val)
        return (
            norm_text_for_compare(adp_label),
            norm_text_for_compare(uzio_label),
            "filing_status: UZIO enum -> ADP label via mapping; punctuation/case/space normalized",
        )
    if ftype == "boolean":
        return to_bool_or_blank(adp_val), to_bool_or_blank(uzio_val), "boolean: Yes/No/1/0/True/False normalized"
    if ftype == "money":
        uzio_is_cents = field in cents_fields
        return (
            to_money_str(adp_val, uzio_is_cents=False),
            to_money_str(uzio_val, uzio_is_cents=uzio_is_cents),
            "money: UZIO ÷100 (cents->dollars)" if uzio_is_cents else "money: numeric compare",
        )
    # default: text normalization
    return norm_text_for_compare(adp_val), norm_text_for_compare(uzio_val), "text: whitespace/case/punctuation normalized"

def is_active_status(val: str) -> bool:
    s = norm_text_for_compare(val)
    # avoids "inactive" being treated as active
    return s == "active" or s.startswith("active ")

# ============================================================
# 7) File reading (Excel/CSV)
# ============================================================
def read_uploaded_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, dtype=str, keep_default_na=False)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        xls = pd.ExcelFile(uploaded)
        sheet = st.selectbox(f"Select sheet: {uploaded.name}", xls.sheet_names, key=f"sheet_{uploaded.name}")
        return pd.read_excel(uploaded, sheet_name=sheet, dtype=str, keep_default_na=False)
    raise ValueError("Unsupported file type. Upload .xlsx/.xls or .csv")

# ============================================================
# 8) Comparison engine
# ============================================================
def compare_data(df_adp: pd.DataFrame, df_uzio: pd.DataFrame, adp_key: str, uzio_key: str, cents_fields: set, compare_all_common: bool) -> Dict[str, pd.DataFrame]:
    df_adp = df_adp.copy()
    df_uzio = df_uzio.copy()

    df_adp[adp_key] = df_adp[adp_key].astype(str).str.strip()
    df_uzio[uzio_key] = df_uzio[uzio_key].astype(str).str.strip()

    adp_idx = df_adp.set_index(adp_key, drop=False)
    uzio_idx = df_uzio.set_index(uzio_key, drop=False)

    keys_adp = set(adp_idx.index) - {""}
    keys_uzio = set(uzio_idx.index) - {""}
    common_keys = sorted(keys_adp & keys_uzio)

    missing_in_uzio = sorted(keys_adp - keys_uzio)
    missing_in_adp = sorted(keys_uzio - keys_adp)

    # Optional descriptive columns
    adp_status_col = pick_first_existing(df_adp, STATUS_CANDIDATES)
    uzio_status_col = pick_first_existing(df_uzio, STATUS_CANDIDATES)

    adp_fn = pick_first_existing(df_adp, NAME_FIRST_CANDIDATES)
    adp_ln = pick_first_existing(df_adp, NAME_LAST_CANDIDATES)
    uzio_fn = pick_first_existing(df_uzio, NAME_FIRST_CANDIDATES)
    uzio_ln = pick_first_existing(df_uzio, NAME_LAST_CANDIDATES)

    # Resolve allowances columns (special computed compare)
    adp_total_col, adp_basic_col, adp_addl_col = resolve_allowance_cols(df_adp)
    uzio_total_col, uzio_basic_col, uzio_addl_col = resolve_allowance_cols(df_uzio)
    allowances_in_scope = any([adp_total_col, adp_basic_col, adp_addl_col, uzio_total_col, uzio_basic_col, uzio_addl_col])

    # Resolve known fields
    resolved_fields = []
    for field, ftype, adp_cands, uzio_cands in FIELD_SPECS:
        adp_col = resolve_column(df_adp, adp_cands)
        uzio_col = resolve_column(df_uzio, uzio_cands)
        if adp_col and uzio_col:
            resolved_fields.append((field, ftype, adp_col, uzio_col))

    # Add relevant common columns (FIT_/SIT_ etc.) to reduce noise
    used_adp_cols = {a for _, _, a, _ in resolved_fields}
    used_uzio_cols = {u for _, _, _, u in resolved_fields}

    common_cols = sorted((set(df_adp.columns) & set(df_uzio.columns)) - {adp_key, uzio_key})
    if compare_all_common:
        extra_cols = common_cols
    else:
        extra_cols = [
            c for c in common_cols
            if c.startswith("FIT_") or c.startswith("SIT_") or c.startswith("FEDERAL_") or c.endswith("_FILING_STATUS")
        ]

    for col in extra_cols:
        if col in used_adp_cols or col in used_uzio_cols:
            continue
        # avoid double comparing allowances if the total column exists as a raw column
        if col in SIT_TOTAL_ALLOWANCES_CANDS or col in SIT_BASIC_ALLOWANCE_CANDS or col in SIT_ADDITIONAL_ALLOWANCES_CANDS:
            continue
        resolved_fields.append((col, "text", col, col))

    mismatches = []

    for k in common_keys:
        adp_row = adp_idx.loc[k]
        uzio_row = uzio_idx.loc[k]

        # status
        status = ""
        if uzio_status_col:
            status = str(uzio_row.get(uzio_status_col, "")).strip()
        elif adp_status_col:
            status = str(adp_row.get(adp_status_col, "")).strip()

        # name
        full_name = ""
        if uzio_fn and uzio_ln:
            full_name = f"{str(uzio_row.get(uzio_fn,'')).strip()} {str(uzio_row.get(uzio_ln,'')).strip()}".strip()
        elif adp_fn and adp_ln:
            full_name = f"{str(adp_row.get(adp_fn,'')).strip()} {str(adp_row.get(adp_ln,'')).strip()}".strip()

        # ----- Special compare: SIT allowances (total OR basic+additional) -----
        if allowances_in_scope:
            adp_allow = compute_sit_total_allowances(adp_row, adp_total_col, adp_basic_col, adp_addl_col)
            uzio_allow = compute_sit_total_allowances(uzio_row, uzio_total_col, uzio_basic_col, uzio_addl_col)

            adp_norm = to_int_allowances(adp_allow) if str(adp_allow).strip() != "" else ""
            uzio_norm = to_int_allowances(uzio_allow) if str(uzio_allow).strip() != "" else ""
            rule = "numeric: SIT allowances; uses SIT_TOTAL_ALLOWANCES else BASIC+ADDITIONAL; True/False->1/0 only for allowances"

            if not (adp_norm == "" and uzio_norm == "") and adp_norm != uzio_norm:
                mismatches.append({
                    "EMPLOYEE_KEY": k,
                    "EMPLOYEE_NAME": full_name,
                    "EMPLOYMENT_STATUS": status,
                    "FIELD": SIT_TOTAL_ALLOWANCES_FIELD,
                    "ADP_COLUMN": allowance_col_label(adp_total_col, adp_basic_col, adp_addl_col),
                    "UZIO_COLUMN": allowance_col_label(uzio_total_col, uzio_basic_col, uzio_addl_col),
                    "ADP_RAW": str(adp_allow),
                    "UZIO_RAW": str(uzio_allow),
                    "ADP_NORMALIZED": adp_norm,
                    "UZIO_NORMALIZED": uzio_norm,
                    "RULE_APPLIED": rule,
                })

        # ----- All other resolved fields -----
        for field, ftype, adp_col, uzio_col in resolved_fields:
            adp_val = adp_row.get(adp_col, "")
            uzio_val = uzio_row.get(uzio_col, "")

            adp_norm, uzio_norm, rule = normalize_pair(field, ftype, adp_val, uzio_val, cents_fields)

            # treat both empty as match
            if (adp_norm == "" or pd.isna(adp_norm)) and (uzio_norm == "" or pd.isna(uzio_norm)):
                continue

            if adp_norm != uzio_norm:
                mismatches.append({
                    "EMPLOYEE_KEY": k,
                    "EMPLOYEE_NAME": full_name,
                    "EMPLOYMENT_STATUS": status,
                    "FIELD": field,
                    "ADP_COLUMN": adp_col,
                    "UZIO_COLUMN": uzio_col,
                    "ADP_RAW": str(adp_val),
                    "UZIO_RAW": str(uzio_val),
                    "ADP_NORMALIZED": adp_norm,
                    "UZIO_NORMALIZED": uzio_norm,
                    "RULE_APPLIED": rule,
                })

    mism_df = pd.DataFrame(mismatches)

    # split active / terminated-ish
    if mism_df.empty:
        active_df = mism_df.copy()
        term_df = mism_df.copy()
    else:
        active_mask = mism_df["EMPLOYMENT_STATUS"].astype(str).map(is_active_status)
        active_df = mism_df[active_mask].copy()
        term_df = mism_df[~active_mask].copy()

    summary = pd.DataFrame([
        {"Metric": "Employees compared (matched keys)", "Value": len(common_keys)},
        {"Metric": "ADP keys missing in UZIO", "Value": len(missing_in_uzio)},
        {"Metric": "UZIO keys missing in ADP", "Value": len(missing_in_adp)},
        {"Metric": "Total mismatches", "Value": int(len(mism_df))},
        {"Metric": "Employees with ≥1 mismatch", "Value": int(mism_df["EMPLOYEE_KEY"].nunique()) if not mism_df.empty else 0},
        {"Metric": "Generated at", "Value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ])

    mismatch_summary = (
        mism_df.groupby("FIELD")
        .agg(MISMATCH_COUNT=("FIELD", "size"), EMPLOYEE_COUNT=("EMPLOYEE_KEY", "nunique"))
        .reset_index()
        .sort_values(["MISMATCH_COUNT", "EMPLOYEE_COUNT"], ascending=False)
        if not mism_df.empty else pd.DataFrame(columns=["FIELD", "MISMATCH_COUNT", "EMPLOYEE_COUNT"])
    )

    field_rules = pd.DataFrame([
        {
            "FIELD": field,
            "TYPE": ftype,
            "ADP_COLUMN": adp_col,
            "UZIO_COLUMN": uzio_col,
            "UZIO_CENTS_DIV100": (field in cents_fields) if ftype == "money" else "",
            "NOTES": ""
        }
        for field, ftype, adp_col, uzio_col in resolved_fields
    ])

    # Add a special row for allowances rule
    if allowances_in_scope:
        field_rules = pd.concat([
            field_rules,
            pd.DataFrame([{
                "FIELD": SIT_TOTAL_ALLOWANCES_FIELD,
                "TYPE": "allowances_int (direct or derived)",
                "ADP_COLUMN": allowance_col_label(adp_total_col, adp_basic_col, adp_addl_col),
                "UZIO_COLUMN": allowance_col_label(uzio_total_col, uzio_basic_col, uzio_addl_col),
                "UZIO_CENTS_DIV100": "",
                "NOTES": "If SIT_TOTAL_ALLOWANCES missing, compare BASIC+ADDITIONAL. Numeric field; True/False->1/0 only here."
            }])
        ], ignore_index=True)

    mapped_uzio_cols = {u for _, _, _, u in resolved_fields} | {uzio_key}
    # also consider allowance cols as "handled"
    mapped_uzio_cols |= set([c for c in [uzio_total_col, uzio_basic_col, uzio_addl_col] if c])

    unverified_uzio = sorted(set(df_uzio.columns) - mapped_uzio_cols)
    unverified_uzio_df = pd.DataFrame({"UZIO_FIELD": unverified_uzio})

    missing_in_uzio_df = pd.DataFrame({"ADP_KEY_MISSING_IN_UZIO": missing_in_uzio})
    missing_in_adp_df = pd.DataFrame({"UZIO_KEY_MISSING_IN_ADP": missing_in_adp})

    return {
        "Summary": summary,
        "Mismatch Summary": mismatch_summary,
        "Mismatches (All)": mism_df,
        "Mismatches (Active)": active_df,
        "Mismatches (Terminated)": term_df,
        "Field Mapping Rules": field_rules,
        "Unverified UZIO Fields": unverified_uzio_df,
        "Missing in UZIO": missing_in_uzio_df,
        "Missing in ADP": missing_in_adp_df,
    }

def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df = df.copy()
            df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            ws.freeze_panes(1, 0)
            # column widths
            for i, col in enumerate(df.columns):
                try:
                    max_len = int(df[col].astype(str).map(len).max()) if len(df) else len(col)
                except Exception:
                    max_len = len(col)
                width = max(12, min(55, max(max_len, len(col)) + 2))
                ws.set_column(i, i, width)
    return output.getvalue()

# ============================================================
# 9) Streamlit UI
# ============================================================
st.set_page_config(page_title="ADP ↔ UZIO FIT/SIT Validator", layout="wide")
st.title("ADP ↔ UZIO FIT/SIT Cross-Verification Utility")

st.markdown(
    """
Uploads: **ADP** + **UZIO** (Excel/CSV)

Rules applied:
- Filing status: UZIO enum → ADP label via your mapping table (then punctuation/case/space normalized compare)
- Yes/No ↔ True/False ↔ 1/0 normalization for boolean fields
- Money compare: UZIO ÷100 for configured cents fields
- **SIT allowances are numeric**:
  - Prefer `SIT_TOTAL_ALLOWANCES`
  - If missing, compare `SIT_BASIC_ALLOWANCE + SIT_ADDITIONAL_ALLOWANCES`
  - If True/False appears, treat as 1/0 **only for allowances**
"""
)

c1, c2 = st.columns(2)
with c1:
    adp_file = st.file_uploader("Upload ADP file (.xlsx/.xls/.csv)", type=["xlsx", "xls", "csv"])
with c2:
    uzio_file = st.file_uploader("Upload UZIO file (.xlsx/.xls/.csv)", type=["xlsx", "xls", "csv"])

if not adp_file or not uzio_file:
    st.stop()

adp_df = read_uploaded_file(adp_file)
uzio_df = read_uploaded_file(uzio_file)

adp_df = standardize_columns(adp_df)
uzio_df = standardize_columns(uzio_df)

st.subheader("Format handling (auto)")
if detect_long_format(adp_df):
    st.info("ADP looks like LONG format (WITHHOLDING_FIELD_KEY/VALUE). It will be pivoted to WIDE.")
    default_ids = [c for c in ["EMPLOYEE_ID", "ASSOCIATE_ID", "FIRST_NAME", "LAST_NAME", "EMPLOYMENT_STATUS"] if c in adp_df.columns]
    id_cols = st.multiselect("ADP pivot ID columns", options=list(adp_df.columns), default=default_ids, key="adp_pivot")
    if not id_cols:
        st.error("Select at least one pivot ID column for ADP long→wide conversion.")
        st.stop()
    adp_df = pivot_long_to_wide(adp_df, id_cols)

if detect_long_format(uzio_df):
    st.info("UZIO looks like LONG format (WITHHOLDING_FIELD_KEY/VALUE). It will be pivoted to WIDE.")
    default_ids = [c for c in ["EMPLOYEE_ID", "ASSOCIATE_ID", "FIRST_NAME", "LAST_NAME", "EMPLOYMENT_STATUS"] if c in uzio_df.columns]
    id_cols = st.multiselect("UZIO pivot ID columns", options=list(uzio_df.columns), default=default_ids, key="uzio_pivot")
    if not id_cols:
        st.error("Select at least one pivot ID column for UZIO long→wide conversion.")
        st.stop()
    uzio_df = pivot_long_to_wide(uzio_df, id_cols)

st.subheader("Key selection (how employees are matched)")
adp_key_guess = pick_first_existing(adp_df, KEY_CANDIDATES) or adp_df.columns[0]
uzio_key_guess = pick_first_existing(uzio_df, KEY_CANDIDATES) or uzio_df.columns[0]

k1, k2 = st.columns(2)
with k1:
    adp_key = st.selectbox("ADP key column", options=list(adp_df.columns), index=list(adp_df.columns).index(adp_key_guess))
with k2:
    uzio_key = st.selectbox("UZIO key column", options=list(uzio_df.columns), index=list(uzio_df.columns).index(uzio_key_guess))

st.subheader("Money fields stored in cents in UZIO (will be divided by 100)")
cents_text = st.text_area("One per line:", value="\n".join(sorted(DEFAULT_CENTS_FIELDS)), height=130)
cents_fields = {normalize_colname(x) for x in cents_text.splitlines() if x.strip()}

compare_all_common = st.checkbox("Compare ALL common columns (can create noisy mismatches)", value=False)

run = st.button("Generate mismatch report", type="primary")

if run:
    sheets = compare_data(adp_df, uzio_df, adp_key, uzio_key, cents_fields, compare_all_common)
    excel_bytes = to_excel_bytes(sheets)

    st.success("Mismatch report generated.")
    st.download_button(
        "Download Excel mismatch report",
        data=excel_bytes,
        file_name=f"ADP_vs_UZIO_FIT_SIT_Mismatch_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Preview: Mismatch Summary")
    st.dataframe(sheets["Mismatch Summary"], use_container_width=True)

    st.subheader("Preview: Mismatches (All) — first 200 rows")
    st.dataframe(sheets["Mismatches (All)"].head(200), use_container_width=True)
