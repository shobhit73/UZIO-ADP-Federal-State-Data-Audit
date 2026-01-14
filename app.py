import io
import re
from datetime import datetime
from typing import Dict, Any, Tuple, Optional, List

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# ============================================================
# 1) UZIO enum -> ADP filing status label mapping (your mapping)
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
# 2) Defaults: UZIO cents fields
# ============================================================
DEFAULT_CENTS_FIELDS = {
    "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
    "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",
    "FIT_DEDUCTIONS_OVER_STANDARD",
    "FIT_OTHER_INCOME",
    "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
}

# ============================================================
# 3) Allowances logic
# ============================================================
SIT_TOTAL_ALLOWANCES_FIELD = "SIT_TOTAL_ALLOWANCES"

SIT_TOTAL_ALLOWANCES_CANDS = ["SIT_TOTAL_ALLOWANCES", "SIT_ALLOWANCES", "STATE_ALLOWANCES"]
SIT_BASIC_ALLOWANCE_CANDS = ["SIT_BASIC_ALLOWANCE", "SIT_BASIC_ALLOWANCES", "STATE_BASIC_ALLOWANCE", "STATE_BASIC_ALLOWANCES", "SIT_ALLOWANCE_BASIC", "SIT_ALLOWANCES_BASIC"]
SIT_ADDITIONAL_ALLOWANCES_CANDS = ["SIT_ADDITIONAL_ALLOWANCES", "SIT_ADDITIONAL_ALLOWANCE", "STATE_ADDITIONAL_ALLOWANCES", "STATE_ADDITIONAL_ALLOWANCE", "SIT_ALLOWANCE_ADDITIONAL", "SIT_ALLOWANCES_ADDITIONAL"]

# ============================================================
# 4) Key + status + name column candidates
# ============================================================
KEY_CANDIDATES = ["EMPLOYEE_ID", "ASSOCIATE_ID", "EE_ID", "EMP_ID", "WORKER_ID", "EMPLOYEEID"]
STATUS_CANDIDATES = ["EMPLOYMENT_STATUS", "STATUS", "EE_STATUS", "EMPLOYEE_STATUS"]
NAME_FIRST_CANDIDATES = ["FIRST_NAME", "EMPLOYEE_FIRST_NAME", "EE_FIRST_NAME", "WORKER_FIRST_NAME"]
NAME_LAST_CANDIDATES = ["LAST_NAME", "EMPLOYEE_LAST_NAME", "EE_LAST_NAME", "WORKER_LAST_NAME"]

# ============================================================
# 5) Boolean + normalize helpers
# ============================================================
YES = {"YES", "Y", "TRUE", "T", "1"}
NO  = {"NO", "N", "FALSE", "F", "0"}

def normalize_colname(c: str) -> str:
    return re.sub(r"\s+", "_", str(c).strip().upper())

def make_unique_columns(cols: List[str]) -> List[str]:
    out = []
    seen = {}
    for c in cols:
        base = c
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}__{seen[base]}")
    return out

def standardize_columns_keep_map(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Standardize columns to UPPER_WITH_UNDERSCORES and keep mapping:
    standardized_name -> original_name
    (If duplicates after standardizing, suffix __2, __3 etc.)
    """
    orig = list(df.columns)
    std = [normalize_colname(c) for c in orig]
    std_unique = make_unique_columns(std)
    df2 = df.copy()
    df2.columns = std_unique
    std_to_orig = {s: o for s, o in zip(std_unique, orig)}
    return df2, std_to_orig

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
    s = re.sub(r"[^a-z0-9]+", " ", s)  # punctuation incl "/" becomes space
    s = re.sub(r"\s+", " ", s).strip()
    return s

def to_bool(v: Any) -> Optional[bool]:
    s = str(v).strip().upper() if v is not None else ""
    if s == "":
        return None
    if s in YES:
        return True
    if s in NO:
        return False
    return None

def to_bool_blank_false(v: Any) -> bool:
    """
    Blank treated as False; Yes/No/True/False/1/0 normalized.
    """
    b = to_bool(v)
    if b is None:
        return False
    return b

def to_int_allowances(v: Any) -> Optional[int]:
    """
    Numeric allowances.
    True/False -> 1/0.
    Blank -> None.
    """
    s = str(v).strip() if v is not None else ""
    if s == "":
        return None
    u = s.upper()
    if u in YES:
        return 1
    if u in NO:
        return 0
    try:
        n = float(s.replace(",", ""))
        return int(n) if n.is_integer() else int(round(n))
    except Exception:
        return None

def to_int_blank_zero(v: Any) -> int:
    """
    Blank -> 0; False -> 0; True -> 1.
    """
    n = to_int_allowances(v)
    return 0 if n is None else n

def to_money(adp_val: Any, uzio_val: Any, uzio_is_cents: bool) -> Tuple[Optional[float], Optional[float]]:
    def parse_money(x: Any, cents: bool) -> Optional[float]:
        s = str(x).strip() if x is not None else ""
        if s == "":
            return None
        u = s.upper()
        if u in YES:
            s = "1"
        if u in NO:
            s = "0"
        try:
            n = float(s.replace(",", ""))
            if cents:
                n = n / 100.0
            return n
        except Exception:
            return None
    return parse_money(adp_val, False), parse_money(uzio_val, uzio_is_cents)

def parse_uzio_enum_to_adp_label(enum_val: Any) -> str:
    raw = str(enum_val).strip() if enum_val is not None else ""
    if raw == "":
        return ""
    u = raw.upper()
    if "_" not in u:
        return raw
    prefix, suffix = u.split("_", 1)
    if prefix in FILING_STATUS_MAP and suffix in FILING_STATUS_MAP[prefix]:
        return FILING_STATUS_MAP[prefix][suffix]
    return suffix.replace("_", " ").title()

def pick_first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def resolve_allowance_cols(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    def first(cands):
        for c in cands:
            if c in df.columns:
                return c
        return None
    return first(SIT_TOTAL_ALLOWANCES_CANDS), first(SIT_BASIC_ALLOWANCE_CANDS), first(SIT_ADDITIONAL_ALLOWANCES_CANDS)

def compute_sit_total_allowances(row: pd.Series, total_col: Optional[str], basic_col: Optional[str], addl_col: Optional[str]) -> Optional[int]:
    """
    Prefer total if present and non-blank, else basic+additional.
    Blank -> None.
    """
    if total_col and total_col in row.index:
        v = to_int_allowances(row.get(total_col, ""))
        if v is not None:
            return v
    has_parts = False
    total = 0
    if basic_col and basic_col in row.index:
        has_parts = True
        total += to_int_blank_zero(row.get(basic_col, ""))
    if addl_col and addl_col in row.index:
        has_parts = True
        total += to_int_blank_zero(row.get(addl_col, ""))
    if not has_parts:
        return None
    return total

def allowance_col_label(total_col: Optional[str], basic_col: Optional[str], addl_col: Optional[str], col_to_orig: Dict[str, str]) -> str:
    if total_col:
        return col_to_orig.get(total_col, total_col)
    parts = [c for c in [basic_col, addl_col] if c]
    if not parts:
        return ""
    return " + ".join([col_to_orig.get(c, c) for c in parts])

def is_active_status(v: Any) -> bool:
    """
    Robust active detection:
    - 'Active', 'A', 'Active Employee', 'Active - ...'
    - Not terminated/inactive
    """
    s = norm_text_for_compare(v)
    if s in {"a", "act"}:
        return True
    if "active" in s and "inactive" not in s and "terminated" not in s and "term" not in s and "separated" not in s:
        return True
    return False

# ============================================================
# 6) “Image-2 style” field definitions (human-friendly)
# Each row: display_field, type_key, uzio_field_name, notes,
#           and ADP keyword matcher function for ADP columns
# ============================================================
class FieldDef:
    def __init__(self, display: str, uzio_field: str, type_key: str, notes: str,
                 adp_keywords_any: List[str], adp_keywords_all: List[str]):
        self.display = display
        self.uzio_field = uzio_field
        self.type_key = type_key
        self.notes = notes
        self.adp_keywords_any = [normalize_colname(x) for x in adp_keywords_any]
        self.adp_keywords_all = [normalize_colname(x) for x in adp_keywords_all]

FIELD_DEFS: List[FieldDef] = [
    FieldDef(
        display="FIT Filing Status",
        uzio_field="FIT_FILING_STATUS",
        type_key="filing_status",
        notes="UZIO enum mapped to ADP label (your mapping table); punctuation/case/space normalized.",
        adp_keywords_any=["MARITAL", "FILING", "STATUS"],
        adp_keywords_all=["FEDERAL", "MARITAL", "STATUS", "DESCRIPTION"],
    ),
    FieldDef(
        display="FIT Multiple Jobs indicator",
        uzio_field="FIT_HIGHER_WITHHOLDING",
        type_key="boolean",
        notes="ADP Yes/No vs UZIO True/False/1/0 normalized.",
        adp_keywords_any=["MULTIPLE", "JOBS"],
        adp_keywords_all=["MULTIPLE", "JOBS", "INDICATOR"],
    ),
    FieldDef(
        display="FIT Dependents amount",
        uzio_field="FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",
        type_key="money_cents",
        notes="UZIO stored in cents; compare ADP dollars vs UZIO cents (÷100). Blank/0 treated as equal.",
        adp_keywords_any=["DEPENDENTS"],
        adp_keywords_all=["DEPENDENTS"],
    ),
    FieldDef(
        display="FIT Deductions amount",
        uzio_field="FIT_DEDUCTIONS_OVER_STANDARD",
        type_key="money_cents",
        notes="UZIO stored in cents; compare ADP dollars vs UZIO cents (÷100). Blank/0 treated as equal.",
        adp_keywords_any=["DEDUCTIONS"],
        adp_keywords_all=["DEDUCTIONS"],
    ),
    FieldDef(
        display="FIT Other income",
        uzio_field="FIT_OTHER_INCOME",
        type_key="money_cents",
        notes="UZIO stored in cents; compare ADP dollars vs UZIO cents (÷100). Blank/0 treated as equal.",
        adp_keywords_any=["OTHER", "INCOME"],
        adp_keywords_all=["OTHER", "INCOME"],
    ),
    FieldDef(
        display="FIT Additional withholding (per pay period)",
        uzio_field="FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
        type_key="money_cents",
        notes="UZIO stored in cents; compare ADP dollars vs UZIO cents (÷100). Blank/0 treated as equal.",
        adp_keywords_any=["ADDITIONAL", "TAX", "AMOUNT"],
        adp_keywords_all=["FEDERAL", "ADDITIONAL", "TAX", "AMOUNT"],
    ),
    FieldDef(
        display="FIT Withholding exemption (Do not calculate FIT)",
        uzio_field="FIT_WITHHOLDING_EXEMPTION",
        type_key="boolean_blank_false",
        notes="ADP blank treated as No/False. ADP Yes/No vs UZIO True/False/1/0 normalized.",
        adp_keywords_any=["DO_NOT_CALCULATE", "FEDERAL", "INCOME", "TAX"],
        adp_keywords_all=["DO", "NOT", "CALCULATE", "FEDERAL", "INCOME", "TAX"],
    ),
    FieldDef(
        display="FIT W-4 exemptions/allowances",
        uzio_field="FIT_WITHHOLDING_ALLOWANCE",
        type_key="int_blank_zero",
        notes="Numeric; UZIO False/blank treated as 0.",
        adp_keywords_any=["W4", "EXEMPTIONS", "ALLOWANCES"],
        adp_keywords_all=["FEDERAL", "W4", "EXEMPTIONS"],
    ),
    FieldDef(
        display="SIT Withholding exemption (Do not calculate SIT)",
        uzio_field="SIT_WITHHOLDING_EXEMPTION",
        type_key="boolean_blank_false",
        notes="ADP blank treated as No/False. ADP Yes/No vs UZIO True/False/1/0 normalized.",
        adp_keywords_any=["DO_NOT_CALCULATE", "STATE", "TAX"],
        adp_keywords_all=["DO", "NOT", "CALCULATE", "STATE", "TAX"],
    ),
    FieldDef(
        display="SIT Total allowances",
        uzio_field="SIT_TOTAL_ALLOWANCES_CALC",
        type_key="allowances_calc",
        notes="Computed: prefer SIT_TOTAL_ALLOWANCES else SIT_BASIC_ALLOWANCE + SIT_ADDITIONAL_ALLOWANCES. Numeric; booleans treated as 0/1.",
        adp_keywords_any=["STATE", "EXEMPTIONS", "ALLOWANCES"],
        adp_keywords_all=["STATE", "EXEMPTIONS", "ALLOWANCES"],
    ),
    FieldDef(
        display="SIT Additional withholding (per pay period)",
        uzio_field="SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
        type_key="money_cents",
        notes="UZIO stored in cents; compare ADP dollars vs UZIO cents (÷100). Blank/0 treated as equal.",
        adp_keywords_any=["STATE", "ADDITIONAL", "TAX", "AMOUNT"],
        adp_keywords_all=["STATE", "ADDITIONAL", "TAX", "AMOUNT"],
    ),
]

# ============================================================
# 7) Column resolver: find ADP column by keyword scoring
# ============================================================
def find_best_col_by_keywords(df_cols: List[str], keywords_any: List[str], keywords_all: List[str]) -> Optional[str]:
    """
    df_cols are standardized column names.
    We score columns by:
      +2 for each keyword in ALL list found as substring
      +1 for each keyword in ANY list found as substring
    Must match at least 1 ANY keyword, and ideally all ALL keywords.
    """
    best = None
    best_score = 0

    for col in df_cols:
        score = 0
        col_u = col

        any_hits = sum(1 for k in keywords_any if k and k in col_u)
        all_hits = sum(1 for k in keywords_all if k and k in col_u)

        if any_hits == 0 and len(keywords_any) > 0:
            continue

        score += any_hits * 1
        score += all_hits * 2

        # bias toward exact-ish columns
        if all_hits == len(keywords_all) and len(keywords_all) > 0:
            score += 3

        if score > best_score:
            best_score = score
            best = col

    return best

def resolve_uzio_col(df: pd.DataFrame, uzio_field: str) -> Optional[str]:
    """
    Prefer exact column match; else try common alternates.
    """
    if uzio_field in df.columns:
        return uzio_field
    # some exports use slightly different naming
    alts = [
        uzio_field.replace("_PER_PAY_PERIOD", ""),
        uzio_field.replace("ADDl", "ADDL"),
        uzio_field.replace("WITHHOLDING", "WITHHOLD"),
    ]
    for a in alts:
        if a in df.columns:
            return a
    return None

# ============================================================
# 8) Compare normalization per type
# ============================================================
def values_equal(type_key: str, adp_raw: Any, uzio_raw: Any, cents_fields: set,
                 allowances_tuple: Tuple[Optional[int], Optional[int]] = (None, None)) -> Tuple[bool, Any, Any]:
    """
    Return (is_equal, adp_norm, uzio_norm) based on type_key
    """
    if type_key == "filing_status":
        adp_norm = norm_text_for_compare(adp_raw)
        uzio_label = parse_uzio_enum_to_adp_label(uzio_raw)
        uzio_norm = norm_text_for_compare(uzio_label)
        return adp_norm == uzio_norm, adp_norm, uzio_norm

    if type_key == "boolean":
        adp_norm = to_bool(adp_raw)
        uzio_norm = to_bool(uzio_raw)
        # treat None==None as equal
        return adp_norm == uzio_norm, adp_norm, uzio_norm

    if type_key == "boolean_blank_false":
        adp_norm = to_bool_blank_false(adp_raw)
        uzio_norm = to_bool_blank_false(uzio_raw)
        return adp_norm == uzio_norm, adp_norm, uzio_norm

    if type_key == "int_blank_zero":
        adp_norm = to_int_blank_zero(adp_raw)
        uzio_norm = to_int_blank_zero(uzio_raw)
        return adp_norm == uzio_norm, adp_norm, uzio_norm

    if type_key == "money_cents":
        # For these fields we assume UZIO cents. ADP dollars.
        adp_n, uzio_n = to_money(adp_raw, uzio_raw, uzio_is_cents=True)
        # blank and 0 treated equal
        a = 0.0 if adp_n is None else float(adp_n)
        u = 0.0 if uzio_n is None else float(uzio_n)
        return abs(a - u) < 0.0001, round(a, 2), round(u, 2)

    if type_key == "allowances_calc":
        adp_norm, uzio_norm = allowances_tuple
        a = 0 if adp_norm is None else int(adp_norm)
        u = 0 if uzio_norm is None else int(uzio_norm)
        return a == u, a, u

    # fallback text
    adp_norm = norm_text_for_compare(adp_raw)
    uzio_norm = norm_text_for_compare(uzio_raw)
    return adp_norm == uzio_norm, adp_norm, uzio_norm

# ============================================================
# 9) Excel output (openpyxl engine)
# ============================================================
def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet_name = name[:31]
            df = df.copy()
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            # widths (scan first 2000 rows max)
            for i, col in enumerate(df.columns, start=1):
                sample = df[col].astype(str).head(2000).tolist()
                max_len = max([len(str(col))] + [len(x) for x in sample if x is not None])
                ws.column_dimensions[get_column_letter(i)].width = min(60, max(12, max_len + 2))
    return output.getvalue()

# ============================================================
# 10) Read file
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
# 11) Main compare
# ============================================================
def compare(adp_df: pd.DataFrame, uzio_df: pd.DataFrame,
            adp_std_to_orig: Dict[str, str], uzio_std_to_orig: Dict[str, str],
            adp_key: str, uzio_key: str, cents_fields: set) -> Dict[str, pd.DataFrame]:

    adp = adp_df.copy()
    uzio = uzio_df.copy()

    # normalize keys
    adp[adp_key] = adp[adp_key].astype(str).str.strip()
    uzio[uzio_key] = uzio[uzio_key].astype(str).str.strip()

    adp_idx = adp.set_index(adp_key, drop=False)
    uzio_idx = uzio.set_index(uzio_key, drop=False)

    adp_keys = set(adp_idx.index) - {""}
    uzio_keys = set(uzio_idx.index) - {""}
    common = sorted(adp_keys & uzio_keys)
    missing_in_uzio = sorted(adp_keys - uzio_keys)
    missing_in_adp = sorted(uzio_keys - adp_keys)

    # status/name columns (for reporting)
    adp_status = pick_first_existing(adp, STATUS_CANDIDATES)
    uzio_status = pick_first_existing(uzio, STATUS_CANDIDATES)
    adp_fn = pick_first_existing(adp, NAME_FIRST_CANDIDATES)
    adp_ln = pick_first_existing(adp, NAME_LAST_CANDIDATES)
    uzio_fn = pick_first_existing(uzio, NAME_FIRST_CANDIDATES)
    uzio_ln = pick_first_existing(uzio, NAME_LAST_CANDIDATES)

    # allowances cols
    adp_total, adp_basic, adp_addl = resolve_allowance_cols(adp)
    uzio_total, uzio_basic, uzio_addl = resolve_allowance_cols(uzio)

    # Build mapping rules table like Image-2
    mapping_rows = []
    resolved_pairs = []  # (FieldDef, adp_col, uzio_col)

    adp_cols = list(adp.columns)

    for f in FIELD_DEFS:
        # ADP column resolve
        adp_col = find_best_col_by_keywords(adp_cols, f.adp_keywords_any, f.adp_keywords_all)

        # UZIO column resolve
        uzio_col = None
        if f.type_key == "allowances_calc":
            uzio_col = "SIT_TOTAL_ALLOWANCES_CALC"
        else:
            uzio_col = resolve_uzio_col(uzio, f.uzio_field)

        mapping_rows.append({
            "Field": f.display,
            "ADP Column": adp_std_to_orig.get(adp_col, "") if adp_col else "",
            "UZIO Field": uzio_col if uzio_col else f.uzio_field,
            "Type": f.type_key,
            "Notes": f.notes,
        })

        if f.type_key == "allowances_calc":
            # allow ADP column missing? still compare if ADP has any allowance column
            resolved_pairs.append((f, adp_col, uzio_col))
        else:
            if adp_col and uzio_col:
                resolved_pairs.append((f, adp_col, uzio_col))

    # mismatches
    mismatches = []
    for k in common:
        adp_row = adp_idx.loc[k]
        uzio_row = uzio_idx.loc[k]

        status = ""
        if uzio_status:
            status = str(uzio_row.get(uzio_status, "")).strip()
        elif adp_status:
            status = str(adp_row.get(adp_status, "")).strip()

        full_name = ""
        if uzio_fn and uzio_ln:
            full_name = f"{str(uzio_row.get(uzio_fn,'')).strip()} {str(uzio_row.get(uzio_ln,'')).strip()}".strip()
        elif adp_fn and adp_ln:
            full_name = f"{str(adp_row.get(adp_fn,'')).strip()} {str(adp_row.get(adp_ln,'')).strip()}".strip()

        # precompute allowances
        adp_allow = compute_sit_total_allowances(adp_row, adp_total, adp_basic, adp_addl)
        uzio_allow = compute_sit_total_allowances(uzio_row, uzio_total, uzio_basic, uzio_addl)

        for f, adp_col, uzio_col in resolved_pairs:
            if f.type_key == "allowances_calc":
                # ADP column shown in mapping tab, but value comes from that column (if found), otherwise fallback to basic logic if exists
                adp_raw = adp_row.get(adp_col, "") if adp_col else ""
                # If ADP total allowances column exists, prefer it; else the computed adp_allow
                adp_val = to_int_allowances(adp_raw)
                if adp_val is None:
                    adp_val = adp_allow
                uzio_val = uzio_allow

                equal, adp_norm, uzio_norm = values_equal("allowances_calc", adp_val, uzio_val, cents_fields, (adp_val, uzio_val))
                if not equal:
                    mismatches.append({
                        "EMPLOYEE_KEY": k,
                        "EMPLOYEE_NAME": full_name,
                        "EMPLOYMENT_STATUS": status,
                        "FIELD": "SIT_TOTAL_ALLOWANCES",
                        "ADP_COLUMN": adp_std_to_orig.get(adp_col, "") if adp_col else allowance_col_label(adp_total, adp_basic, adp_addl, adp_std_to_orig),
                        "UZIO_COLUMN": allowance_col_label(uzio_total, uzio_basic, uzio_addl, uzio_std_to_orig),
                        "ADP_RAW": str(adp_val if adp_val is not None else ""),
                        "UZIO_RAW": str(uzio_val if uzio_val is not None else ""),
                        "ADP_NORMALIZED": adp_norm,
                        "UZIO_NORMALIZED": uzio_norm,
                        "RULE_APPLIED": f.notes,
                    })
                continue

            adp_raw = adp_row.get(adp_col, "")
            uzio_raw = uzio_row.get(uzio_col, "")

            equal, adp_norm, uzio_norm = values_equal(f.type_key, adp_raw, uzio_raw, cents_fields)

            # treat both blank-ish equal for text/filing_status; for others already handled as blank->0 or blank->false
            if f.type_key in {"filing_status"}:
                if adp_norm == "" and uzio_norm == "":
                    continue

            if not equal:
                mismatches.append({
                    "EMPLOYEE_KEY": k,
                    "EMPLOYEE_NAME": full_name,
                    "EMPLOYMENT_STATUS": status,
                    "FIELD": f.uzio_field,
                    "ADP_COLUMN": adp_std_to_orig.get(adp_col, adp_col),
                    "UZIO_COLUMN": uzio_col,
                    "ADP_RAW": str(adp_raw),
                    "UZIO_RAW": str(uzio_raw),
                    "ADP_NORMALIZED": adp_norm,
                    "UZIO_NORMALIZED": uzio_norm,
                    "RULE_APPLIED": f.notes,
                })

    mism_df = pd.DataFrame(mismatches)

    if mism_df.empty:
        active_df = mism_df.copy()
        term_df = mism_df.copy()
    else:
        active_mask = mism_df["EMPLOYMENT_STATUS"].map(is_active_status)
        active_df = mism_df[active_mask].copy()
        term_df = mism_df[~active_mask].copy()

    summary = pd.DataFrame([
        {"Metric": "Employees compared (matched keys)", "Value": len(common)},
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

    field_mapping_rules = pd.DataFrame(mapping_rows)

    # Unverified UZIO fields (not used by our mapping)
    used_uzio = set()
    for f, _, uzio_col in resolved_pairs:
        if uzio_col:
            used_uzio.add(uzio_col)
    # also include allowance-related cols
    used_uzio |= {c for c in [uzio_total, uzio_basic, uzio_addl] if c}
    used_uzio |= {uzio_key}

    unverified = sorted(set(uzio.columns) - used_uzio)
    unverified_df = pd.DataFrame({"UZIO_FIELD": [uzio_std_to_orig.get(c, c) for c in unverified]})

    missing_in_uzio_df = pd.DataFrame({"ADP_KEY_MISSING_IN_UZIO": missing_in_uzio})
    missing_in_adp_df = pd.DataFrame({"UZIO_KEY_MISSING_IN_ADP": missing_in_adp})

    return {
        "Summary": summary,
        "Mismatch Summary": mismatch_summary,
        "Mismatches (All)": mism_df,
        "Mismatches (Active)": active_df,
        "Mismatches (Terminated)": term_df,
        "Field Mapping Rules": field_mapping_rules,
        "Unverified UZIO Fields": unverified_df,
        "Missing in UZIO": missing_in_uzio_df,
        "Missing in ADP": missing_in_adp_df,
    }

# ============================================================
# 12) Streamlit UI
# ============================================================
st.set_page_config(page_title="ADP ↔ UZIO FIT/SIT Validator", layout="wide")
st.title("ADP ↔ UZIO FIT/SIT Cross-Verification Utility")

st.markdown(
    """
This utility generates the same style report you expected (like your earlier “Image 2” output):
- Strong ADP column detection (works with headers like “Federal/W4 Marital Status Description”)
- Robust Active/Terminated split (supports A/T, Active Employee, etc.)
- SIT allowances: uses total OR (basic + additional)
- UZIO cents to dollars conversion for configured money fields
"""
)

c1, c2 = st.columns(2)
with c1:
    adp_file = st.file_uploader("Upload ADP file (.xlsx/.xls/.csv)", type=["xlsx", "xls", "csv"])
with c2:
    uzio_file = st.file_uploader("Upload UZIO file (.xlsx/.xls/.csv)", type=["xlsx", "xls", "csv"])

if not adp_file or not uzio_file:
    st.stop()

raw_adp = read_uploaded_file(adp_file)
raw_uzio = read_uploaded_file(uzio_file)

# standardize but keep original header mappings
adp_df, adp_std_to_orig = standardize_columns_keep_map(raw_adp)
uzio_df, uzio_std_to_orig = standardize_columns_keep_map(raw_uzio)

st.subheader("Format handling (auto)")
if detect_long_format(adp_df):
    st.info("ADP looks like LONG format (WITHHOLDING_FIELD_KEY/VALUE). It will be pivoted to WIDE.")
    default_ids = [c for c in ["EMPLOYEE_ID", "ASSOCIATE_ID", "FIRST_NAME", "LAST_NAME", "EMPLOYMENT_STATUS"] if c in adp_df.columns]
    id_cols = st.multiselect("ADP pivot ID columns", options=list(adp_df.columns), default=default_ids, key="adp_pivot")
    if not id_cols:
        st.error("Select at least one pivot ID column for ADP long→wide conversion.")
        st.stop()
    adp_df = pivot_long_to_wide(adp_df, id_cols)
    adp_std_to_orig = {c: c for c in adp_df.columns}  # pivot produces synthetic columns

if detect_long_format(uzio_df):
    st.info("UZIO looks like LONG format (WITHHOLDING_FIELD_KEY/VALUE). It will be pivoted to WIDE.")
    default_ids = [c for c in ["EMPLOYEE_ID", "ASSOCIATE_ID", "FIRST_NAME", "LAST_NAME", "EMPLOYMENT_STATUS"] if c in uzio_df.columns]
    id_cols = st.multiselect("UZIO pivot ID columns", options=list(uzio_df.columns), default=default_ids, key="uzio_pivot")
    if not id_cols:
        st.error("Select at least one pivot ID column for UZIO long→wide conversion.")
        st.stop()
    uzio_df = pivot_long_to_wide(uzio_df, id_cols)
    uzio_std_to_orig = {c: c for c in uzio_df.columns}

st.subheader("Key selection (how employees are matched)")
adp_key_guess = pick_first_existing(adp_df, KEY_CANDIDATES) or adp_df.columns[0]
uzio_key_guess = pick_first_existing(uzio_df, KEY_CANDIDATES) or uzio_df.columns[0]

k1, k2 = st.columns(2)
with k1:
    adp_key = st.selectbox("ADP key column", options=list(adp_df.columns), index=list(adp_df.columns).index(adp_key_guess))
with k2:
    uzio_key = st.selectbox("UZIO key column", options=list(uzio_df.columns), index=list(uzio_df.columns).index(uzio_key_guess))

st.subheader("Money fields stored in cents in UZIO (÷100)")
cents_text = st.text_area("One per line:", value="\n".join(sorted(DEFAULT_CENTS_FIELDS)), height=130)
cents_fields = {normalize_colname(x) for x in cents_text.splitlines() if x.strip()}

run = st.button("Generate mismatch report", type="primary")

if run:
    sheets = compare(adp_df, uzio_df, adp_std_to_orig, uzio_std_to_orig, adp_key, uzio_key, cents_fields)
    excel_bytes = to_excel_bytes(sheets)

    st.success("Mismatch report generated.")
    st.download_button(
        "Download Excel mismatch report",
        data=excel_bytes,
        file_name=f"ADP_vs_UZIO_FIT_SIT_Mismatch_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Preview: Field Mapping Rules")
    st.dataframe(sheets["Field Mapping Rules"], use_container_width=True)

    st.subheader("Preview: Mismatch Summary")
    st.dataframe(sheets["Mismatch Summary"], use_container_width=True)

    st.subheader("Preview: Mismatches (Active)")
    st.dataframe(sheets["Mismatches (Active)"].head(200), use_container_width=True)
