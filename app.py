# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "ADP vs UZIO FIT/SIT Mismatch Audit Tool (Simple)"

# -----------------------------
# UI
# -----------------------------
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
    unsafe_allow_html=True,
)
st.title(APP_TITLE)
st.write(
    "Upload the **ADP** Excel export and the **UZIO** Fed/State withholding Excel export. "
    "Output has 3 tabs: Summary, Field Summary, Comparison Detail."
)

# -----------------------------
# Helpers
# -----------------------------
def norm_blank(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    if isinstance(x, str) and x.strip().lower() in {"", "nan", "none", "null"}:
        return ""
    return x

def norm_colname(c: str) -> str:
    if c is None:
        return ""
    s = str(c).replace("\n", " ").replace("\r", " ").replace("\u00A0", " ")
    s = s.replace("’", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"\s+", " ", s).strip()
    return s

def std_col(c: str) -> str:
    s = norm_colname(c).upper()
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return s

def find_first_existing_col(df: pd.DataFrame, *cands) -> str | None:
    norm_map = {std_col(c): c for c in df.columns}
    for c in cands:
        k = std_col(c)
        if k in norm_map:
            return norm_map[k]
    return None

def safe_str(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (np.integer, int)):
        return str(int(x))
    if isinstance(x, (np.floating, float)):
        if float(x).is_integer():
            return str(int(x))
        return str(x)
    return str(x)

def cf(x) -> str:
    return safe_str(x).strip().casefold()

def try_parse_date(x):
    x = norm_blank(x)
    if x == "":
        return None
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x, errors="coerce")
    s = str(x).strip()
    if s == "" or s in {"00/00/0000", "0/0/0000", "0000-00-00"}:
        return None
    return pd.to_datetime(s, errors="coerce")

def norm_bool_blank_false(x) -> str:
    s = cf(x)
    if s in {"", "0", "false", "f", "no", "n", "off"}:
        return "false"
    if s in {"1", "true", "t", "yes", "y", "on"}:
        return "true"
    if isinstance(x, bool):
        return "true" if x else "false"
    return s  # fallback

def norm_int_blank_zero(x) -> str:
    x = norm_blank(x)
    if x == "" or x is False:
        return "0"
    if x is True:
        return "1"
    try:
        v = float(str(x).replace(",", "").strip())
        return str(int(round(v)))
    except Exception:
        s = re.sub(r"[^\d\-]", "", str(x))
        return s if s not in {"", "-"} else "0"

def norm_money_adp_dollars(x) -> str:
    """ADP values are usually dollars; treat blank/False as 0.00"""
    x = norm_blank(x)
    if x == "" or x is False:
        return "0.00"
    if x is True:
        return "1.00"
    try:
        v = float(str(x).replace(",", "").strip())
        return f"{v:.2f}"
    except Exception:
        return "0.00"

def norm_money_uzio_cents_to_dollars(x) -> str:
    """UZIO values are usually cents; treat blank/False as 0.00"""
    x = norm_blank(x)
    if x == "" or x is False:
        return "0.00"
    if x is True:
        return "0.01"
    try:
        v = float(str(x).replace(",", "").strip())
        return f"{(v / 100.0):.2f}"
    except Exception:
        return "0.00"

def parse_filing_status_map_from_text(text: str) -> dict:
    """
    Parses lines like:
      FEDERAL_SINGLE("Single"), FEDERAL_MARRIED_JOINTLY("Married filing jointly ...")
    into { "FEDERAL_SINGLE": "Single", ... }
    """
    out = {}
    for enum_key, label in re.findall(r'([A-Z0-9_]+)\("([^"]*)"\)', text):
        out[enum_key.strip()] = label.strip()
    return out

def load_default_mapping() -> pd.DataFrame:
    """Minimal mapping (same spirit as your simple census utility)."""
    rows = [
        ("employee_id", "Associate ID"),
        ("employee_first_name", "Legal First Name"),
        ("employee_last_name", "Legal Last Name"),

        ("FIT_FILING_STATUS", "Federal/W4 Marital Status Description"),
        ("FIT_WITHHOLDING_ALLOWANCE", "Federal/W4 Exemptions"),
        ("FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "Federal Additional Tax Amount"),
        ("FIT_WITHHOLDING_EXEMPTION", "Do Not Calculate Federal Income Tax"),
        ("FIT_WITHHOLD_AS_NON_RESIDENT", "Non-Resident Alien"),
        ("FIT_HIGHER_WITHHOLDING", "Multiple Jobs indicator"),
        ("FIT_CHILD_AND_DEPENDENT_TAX_CREDIT", "Dependents"),
        ("FIT_DEDUCTIONS_OVER_STANDARD", "Deductions"),
        ("FIT_OTHER_INCOME", "Other Income"),

        ("SIT_WITHHOLDING_EXEMPTION", "Do not calculate State Tax"),
        ("SIT_FILING_STATUS", "State Marital Status Description"),
        ("SIT_TOTAL_ALLOWANCES", "State Exemptions/Allowances"),
        ("SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD", "State Additional Tax Amount"),
    ]
    return pd.DataFrame(rows, columns=["UZIO_Field", "ADP_Column"])

def resolve_mapping_sheet() -> pd.DataFrame:
    """
    If a local mapping file exists next to the app, use it.
    Else use default mapping above.
    Expected local file:
      Mapping Sheet Data of UZIO and ADP.xlsx
    """
    try:
        mp = pd.ExcelFile("Mapping Sheet Data of UZIO and ADP.xlsx")
        df = pd.read_excel(mp, sheet_name=mp.sheet_names[0], dtype=object)
        df.columns = [norm_colname(c) for c in df.columns]
        ucol = find_first_existing_col(df, "Uzio Columns")
        acol = find_first_existing_col(df, "ADP Columns")
        if ucol and acol:
            out = df[[ucol, acol]].rename(columns={ucol: "UZIO_Field", acol: "ADP_Column"}).copy()
            out["UZIO_Field"] = out["UZIO_Field"].map(norm_colname)
            out["ADP_Column"] = out["ADP_Column"].map(norm_colname)
            out = out.dropna()
            out = out[(out["UZIO_Field"] != "") & (out["ADP_Column"] != "")]
            return out.reset_index(drop=True)
    except Exception:
        pass
    return load_default_mapping()

def resolve_filing_status_map() -> dict:
    """
    Loads from local: filing status_code.txt
    Falls back to a minimal map if not found.
    """
    try:
        with open("filing status_code.txt", "r", encoding="utf-8") as f:
            txt = f.read()
        m = parse_filing_status_map_from_text(txt)
        if m:
            return m
    except Exception:
        pass

    # Fallback (minimal)
    return {
        "FEDERAL_SINGLE": "Single",
        "FEDERAL_MARRIED": "Married",
        "FEDERAL_MARRIED_JOINTLY": "Married filing jointly or Qualifying surviving spouse",
        "FEDERAL_HEAD_OF_HOUSEHOLD": "Head of household",
        "FEDERAL_SINGLE_OR_MARRIED": "Single or Married filing separately",
    }

def pick_latest_adp_row(adp_emp_df: pd.DataFrame) -> pd.Series:
    """Pick ADP row with latest Federal/W4 Effective Date; if missing, take first."""
    if adp_emp_df.empty:
        return pd.Series(dtype=object)
    if "_FED_EFF_DT" in adp_emp_df.columns and adp_emp_df["_FED_EFF_DT"].notna().any():
        idx = adp_emp_df["_FED_EFF_DT"].idxmax()
        return adp_emp_df.loc[idx]
    return adp_emp_df.iloc[0]

# -----------------------------
# Normalization per field
# -----------------------------
MONEY_FIELDS = {
    "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",
    "FIT_DEDUCTIONS_OVER_STANDARD",
    "FIT_OTHER_INCOME",
    "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
    "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",
}
INT_FIELDS = {"FIT_WITHHOLDING_ALLOWANCE", "SIT_TOTAL_ALLOWANCES"}
BOOL_FIELDS = {
    "FIT_WITHHOLDING_EXEMPTION",
    "FIT_WITHHOLD_AS_NON_RESIDENT",
    "FIT_HIGHER_WITHHOLDING",
    "SIT_WITHHOLDING_EXEMPTION",
}
FILING_FIELDS = {"FIT_FILING_STATUS", "SIT_FILING_STATUS"}

def filing_status_token_from_adp(adp_val: str) -> str:
    """
    Normalize ADP filing status to stable tokens.
    Specifically handle:
      "Married filing jointly or Qualifying surviving spouse"
    """
    s = cf(adp_val)
    if s == "":
        return ""
    # tolerant matching
    if "married filing jointly" in s:
        # ADP can be exactly the long value or shorter; both should map to MFJ
        # The user rule: if UZIO is FEDERAL_MARRIED_JOINTLY and ADP is the long label, it's a match.
        return "MFJ_QSS"
    if "qualifying surviving spouse" in s:
        return "MFJ_QSS"
    if "head of household" in s:
        return "HOH"
    if "single" in s and "married" not in s:
        return "SINGLE"
    if "married filing separately" in s:
        return "MFS"
    if "married" in s:
        return "MARRIED"
    return s  # fallback

def filing_status_token_from_uzio(uzio_raw, filing_status_map: dict) -> str:
    """
    Convert UZIO enum/code to a token, using the code itself plus map label as needed.
    Includes your special rule:
      UZIO = FEDERAL_MARRIED_JOINTLY should match ADP "Married filing jointly or Qualifying surviving spouse"
    """
    raw = safe_str(uzio_raw).strip()
    if raw == "":
        return ""
    raw_uc = raw.upper()

    # Explicit rule requested
    if raw_uc == "FEDERAL_MARRIED_JOINTLY":
        return "MFJ_QSS"

    # Try to map enum -> label and then tokenize like ADP
    label = filing_status_map.get(raw_uc) or filing_status_map.get(raw) or raw
    # Tokenize based on mapped label (and still tolerant)
    return filing_status_token_from_adp(label)

def normalize_compare(field: str, uzio_val, adp_val, filing_status_map: dict):
    """
    Returns (uz_norm, adp_norm) strings
    """
    field = norm_colname(field)

    if field in FILING_FIELDS:
        # --- FIX HERE: special handling for FEDERAL_MARRIED_JOINTLY vs ADP long label ---
        uz_token = filing_status_token_from_uzio(uzio_val, filing_status_map)
        adp_token = filing_status_token_from_adp(adp_val)
        return (uz_token, adp_token)

    if field in BOOL_FIELDS:
        return (norm_bool_blank_false(uzio_val), norm_bool_blank_false(adp_val))

    if field in INT_FIELDS:
        return (norm_int_blank_zero(uzio_val), norm_int_blank_zero(adp_val))

    if field in MONEY_FIELDS:
        # UZIO cents -> dollars; ADP dollars -> dollars
        return (norm_money_uzio_cents_to_dollars(uzio_val), norm_money_adp_dollars(adp_val))

    # default: casefold string compare
    return (cf(uzio_val), cf(adp_val))

# -----------------------------
# Core runner: returns Excel bytes
# -----------------------------
def run_audit(adp_bytes: bytes, uzio_bytes: bytes) -> bytes:
    mapping = resolve_mapping_sheet()
    filing_status_map = resolve_filing_status_map()

    # --- Read ADP ---
    adp_xls = pd.ExcelFile(io.BytesIO(adp_bytes), engine="openpyxl")
    adp_sheet = "Data" if "Data" in adp_xls.sheet_names else adp_xls.sheet_names[0]
    adp = pd.read_excel(adp_xls, sheet_name=adp_sheet, dtype=object)
    adp.columns = [norm_colname(c) for c in adp.columns]

    ADP_ID = find_first_existing_col(adp, "Associate ID", "AssociateID", "Employee ID", "EmployeeID")
    if ADP_ID is None:
        raise ValueError("ADP key column not found (expected 'Associate ID' or similar).")

    # effective date for latest selection
    fed_eff_col = find_first_existing_col(adp, "Federal/W4 Effective Date")
    adp["_FED_EFF_DT"] = adp[fed_eff_col].map(try_parse_date) if fed_eff_col else pd.NaT

    ADP_STATE = find_first_existing_col(adp, "State Tax Code", "Lived In State Tax Code")

    # --- Read UZIO (long format) ---
    uz_xls = pd.ExcelFile(io.BytesIO(uzio_bytes), engine="openpyxl")
    uz_sheet = uz_xls.sheet_names[0]
    uz = pd.read_excel(uz_xls, sheet_name=uz_sheet, dtype=object)
    uz.columns = [norm_colname(c) for c in uz.columns]

    UZ_ID = find_first_existing_col(uz, "employee_id", "Employee ID", "EmployeeID")
    if UZ_ID is None:
        raise ValueError("UZIO key column not found (expected 'employee_id').")

    UZ_FN = find_first_existing_col(uz, "employee_first_name", "first_name", "Employee First Name")
    UZ_LN = find_first_existing_col(uz, "employee_last_name", "last_name", "Employee Last Name")
    UZ_STATUS = find_first_existing_col(uz, "status", "employment_status", "Employment Status")
    UZ_SCOPE = find_first_existing_col(uz, "tax_scope")
    UZ_STATE = find_first_existing_col(uz, "state_code")
    UZ_KEY = find_first_existing_col(uz, "withholding_field_key")
    UZ_VAL = find_first_existing_col(uz, "withholding_field_value")
    UZ_EFF = find_first_existing_col(uz, "effective_date")

    for need, name in [(UZ_KEY, "withholding_field_key"), (UZ_VAL, "withholding_field_value")]:
        if need is None:
            raise ValueError(f"UZIO column missing: '{name}'")

    uz["_EFF_DT"] = uz[UZ_EFF].map(try_parse_date) if UZ_EFF else pd.NaT

    # --- Build UZIO latest-per-field lookups ---
    meta_cols = [c for c in [UZ_ID, UZ_FN, UZ_LN, UZ_STATUS] if c]
    uz_meta = uz.groupby(UZ_ID, sort=False).head(1)[meta_cols].copy()
    uz_meta = uz_meta.set_index(UZ_ID)

    # Federal
    uz_fed = uz.copy()
    if UZ_SCOPE and UZ_SCOPE in uz_fed.columns:
        uz_fed = uz_fed[uz_fed[UZ_SCOPE].astype(str).str.upper().eq("FEDERAL")]
    uz_fed = uz_fed.sort_values("_EFF_DT").groupby([UZ_ID, UZ_KEY], as_index=False).tail(1)
    uz_fed_lookup = {(str(r[UZ_ID]), str(r[UZ_KEY])): r[UZ_VAL] for _, r in uz_fed.iterrows()}

    # State
    uz_state = uz.copy()
    if UZ_SCOPE and UZ_SCOPE in uz_state.columns:
        uz_state = uz_state[uz_state[UZ_SCOPE].astype(str).str.upper().eq("STATE")]
    if UZ_STATE and UZ_STATE in uz_state.columns:
        uz_state["_STATE_CODE"] = uz_state[UZ_STATE].map(lambda x: safe_str(x).strip().upper())
    else:
        uz_state["_STATE_CODE"] = ""
    uz_state = uz_state.sort_values("_EFF_DT").groupby([UZ_ID, "_STATE_CODE", UZ_KEY], as_index=False).tail(1)
    uz_state_lookup = {(str(r[UZ_ID]), str(r["_STATE_CODE"]), str(r[UZ_KEY])): r[UZ_VAL] for _, r in uz_state.iterrows()}

    # --- Build ADP latest lookups ---
    adp[ADP_ID] = adp[ADP_ID].map(lambda x: safe_str(x).strip())
    adp_emp_latest = adp.groupby(ADP_ID, sort=False).apply(pick_latest_adp_row).reset_index(drop=True)

    # State rows latest per (emp, state_code)
    if ADP_STATE and ADP_STATE in adp.columns:
        tmp = adp.copy()
        tmp["_STATE_CODE"] = tmp[ADP_STATE].map(lambda x: safe_str(x).strip().upper())
        tmp = tmp.sort_values("_FED_EFF_DT")
        adp_state_latest = tmp.groupby([ADP_ID, "_STATE_CODE"], as_index=False).tail(1)
        adp_state_latest = adp_state_latest.set_index([ADP_ID, "_STATE_CODE"])
    else:
        adp_state_latest = pd.DataFrame().set_index(
            pd.MultiIndex.from_arrays([[], []], names=[ADP_ID, "_STATE_CODE"])
        )

    adp_emp_latest = adp_emp_latest.set_index(ADP_ID)

    # --- Determine employee/state universe ---
    uz_emp_ids = set(map(str, uz_meta.index.astype(str)))
    adp_emp_ids = set(map(str, adp_emp_latest.index.astype(str)))
    all_emp_ids = sorted(uz_emp_ids | adp_emp_ids)

    uz_states = set()
    if len(uz_state):
        uz_states = set(
            (str(r[UZ_ID]), str(r["_STATE_CODE"]))
            for _, r in uz_state[[UZ_ID, "_STATE_CODE"]].drop_duplicates().iterrows()
        )
    adp_states = set()
    if not adp_state_latest.empty:
        adp_states = set((str(i[0]), str(i[1])) for i in adp_state_latest.index)
    all_emp_state = sorted(uz_states | adp_states)

    # --- Compare ---
    rows = []
    for _, m in mapping.iterrows():
        uz_field = norm_colname(m["UZIO_Field"])
        adp_col = norm_colname(m["ADP_Column"])

        # Employee-level fields
        if uz_field in {"employee_id", "employee_first_name", "employee_last_name"}:
            for emp_id in all_emp_ids:
                uz_exists = emp_id in uz_meta.index.astype(str)
                adp_exists = emp_id in adp_emp_latest.index.astype(str)

                uz_val = ""
                if uz_exists and uz_field in {"employee_first_name", "employee_last_name"}:
                    col = UZ_FN if uz_field == "employee_first_name" else UZ_LN
                    if col and col in uz_meta.columns and emp_id in uz_meta.index:
                        uz_val = uz_meta.loc[emp_id, col]
                elif uz_field == "employee_id":
                    uz_val = emp_id if uz_exists else ""

                adp_val = ""
                if adp_exists and adp_col in adp_emp_latest.columns and emp_id in adp_emp_latest.index:
                    adp_val = adp_emp_latest.loc[emp_id, adp_col]

                emp_status = ""
                if UZ_STATUS and uz_exists and UZ_STATUS in uz_meta.columns and emp_id in uz_meta.index:
                    emp_status = uz_meta.loc[emp_id, UZ_STATUS]

                if uz_exists and not adp_exists:
                    status = "MISSING_IN_ADP"
                elif adp_exists and not uz_exists:
                    status = "MISSING_IN_UZIO"
                elif adp_col not in adp.columns:
                    status = "ADP_COLUMN_MISSING"
                else:
                    uz_n, adp_n = normalize_compare(uz_field, uz_val, adp_val, filing_status_map)
                    if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                        status = "OK"
                    elif uz_n == "" and adp_n != "":
                        status = "UZIO_MISSING_VALUE"
                    elif uz_n != "" and adp_n == "":
                        status = "ADP_MISSING_VALUE"
                    else:
                        status = "MISMATCH"

                rows.append(
                    {
                        "Employee ID": emp_id,
                        "State Code": "",
                        "Tax Scope": "EMPLOYEE",
                        "Field": uz_field,
                        "Employment Status": emp_status,
                        "UZIO_Value": uz_val,
                        "ADP_Value": adp_val,
                        "Status": status,
                    }
                )
            continue

        # Federal fields
        if uz_field.startswith("FIT_"):
            for emp_id in all_emp_ids:
                uz_exists = emp_id in uz_emp_ids
                adp_exists = emp_id in adp_emp_ids

                emp_status = ""
                if UZ_STATUS and uz_exists and emp_id in uz_meta.index and UZ_STATUS in uz_meta.columns:
                    emp_status = uz_meta.loc[emp_id, UZ_STATUS]

                uz_val = uz_fed_lookup.get((emp_id, uz_field), "") if uz_exists else ""
                adp_val = (
                    adp_emp_latest.loc[emp_id, adp_col]
                    if (adp_exists and adp_col in adp_emp_latest.columns and emp_id in adp_emp_latest.index)
                    else ""
                )

                if uz_exists and not adp_exists:
                    status = "MISSING_IN_ADP"
                elif adp_exists and not uz_exists:
                    status = "MISSING_IN_UZIO"
                elif adp_col not in adp.columns:
                    status = "ADP_COLUMN_MISSING"
                else:
                    uz_n, adp_n = normalize_compare(uz_field, uz_val, adp_val, filing_status_map)
                    if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                        status = "OK"
                    elif uz_n == "" and adp_n != "":
                        status = "UZIO_MISSING_VALUE"
                    elif uz_n != "" and adp_n == "":
                        status = "ADP_MISSING_VALUE"
                    else:
                        status = "MISMATCH"

                rows.append(
                    {
                        "Employee ID": emp_id,
                        "State Code": "",
                        "Tax Scope": "FED",
                        "Field": uz_field,
                        "Employment Status": emp_status,
                        "UZIO_Value": uz_val,
                        "ADP_Value": adp_val,
                        "Status": status,
                    }
                )
            continue

        # State fields
        if uz_field.startswith("SIT_"):
            for (emp_id, state_code) in all_emp_state:
                uz_exists = emp_id in uz_emp_ids
                adp_exists = emp_id in adp_emp_ids

                emp_status = ""
                if UZ_STATUS and uz_exists and emp_id in uz_meta.index and UZ_STATUS in uz_meta.columns:
                    emp_status = uz_meta.loc[emp_id, UZ_STATUS]

                uz_val = uz_state_lookup.get((emp_id, state_code, uz_field), "") if uz_exists else ""

                adp_val = ""
                if (
                    not adp_state_latest.empty
                    and (emp_id, state_code) in adp_state_latest.index
                    and adp_col in adp_state_latest.columns
                ):
                    adp_val = adp_state_latest.loc[(emp_id, state_code), adp_col]

                if uz_exists and not adp_exists:
                    status = "MISSING_IN_ADP"
                elif adp_exists and not uz_exists:
                    status = "MISSING_IN_UZIO"
                elif adp_col not in adp.columns:
                    status = "ADP_COLUMN_MISSING"
                else:
                    uz_n, adp_n = normalize_compare(uz_field, uz_val, adp_val, filing_status_map)
                    if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                        status = "OK"
                    elif uz_n == "" and adp_n != "":
                        status = "UZIO_MISSING_VALUE"
                    elif uz_n != "" and adp_n == "":
                        status = "ADP_MISSING_VALUE"
                    else:
                        status = "MISMATCH"

                rows.append(
                    {
                        "Employee ID": emp_id,
                        "State Code": state_code,
                        "Tax Scope": "SIT",
                        "Field": uz_field,
                        "Employment Status": emp_status,
                        "UZIO_Value": uz_val,
                        "ADP_Value": adp_val,
                        "Status": status,
                    }
                )

    comparison_detail = pd.DataFrame(
        rows,
        columns=[
            "Employee ID",
            "State Code",
            "Tax Scope",
            "Field",
            "Employment Status",
            "UZIO_Value",
            "ADP_Value",
            "Status",
        ],
    )

    # Field Summary
    statuses = [
        "OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "ADP_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_ADP",
        "ADP_COLUMN_MISSING",
    ]
    if len(comparison_detail):
        field_summary = (
            comparison_detail.pivot_table(
                index=["Tax Scope", "Field"],
                columns="Status",
                values="Employee ID",
                aggfunc="count",
                fill_value=0,
            )
            .reset_index()
        )
        for c in statuses:
            if c not in field_summary.columns:
                field_summary[c] = 0
        field_summary["Total"] = field_summary[statuses].sum(axis=1)
        field_summary = field_summary[["Tax Scope", "Field"] + statuses + ["Total"]]
    else:
        field_summary = pd.DataFrame(columns=["Tax Scope", "Field"] + statuses + ["Total"])

    # Summary
    mismatch_rows = comparison_detail[comparison_detail["Status"].ne("OK")]
    active_mask = mismatch_rows["Employment Status"].astype(str).str.upper().str.contains("ACTIVE", na=False)
    terminated_mask = mismatch_rows["Employment Status"].astype(str).str.upper().str.contains("TERM", na=False)

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total ADP Employees",
                "Total UZIO Employees",
                "Employees in both",
                "Employees only in ADP",
                "Employees only in UZIO",
                "Total Comparisons (field-level rows)",
                "Total Mismatch Rows (non-OK)",
                "Mismatch Rows (Active)",
                "Mismatch Rows (Terminated)",
                "Employees with ≥1 mismatch",
            ],
            "Value": [
                len(adp_emp_ids),
                len(uz_emp_ids),
                len(adp_emp_ids & uz_emp_ids),
                len(adp_emp_ids - uz_emp_ids),
                len(uz_emp_ids - adp_emp_ids),
                int(len(comparison_detail)),
                int(len(mismatch_rows)),
                int(active_mask.sum()),
                int(terminated_mask.sum()),
                int(mismatch_rows["Employee ID"].nunique()),
            ],
        }
    )

    # Write 3 tabs
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary.to_excel(writer, sheet_name="Field Summary", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison Detail", index=False)

    return out.getvalue()

# -----------------------------
# Streamlit controls
# -----------------------------
adp_file = st.file_uploader("Upload ADP Excel (.xlsx)", type=["xlsx"], key="adp")
uzio_file = st.file_uploader("Upload UZIO Excel (.xlsx)", type=["xlsx"], key="uzio")

run_btn = st.button("Run Audit", type="primary", disabled=(adp_file is None or uzio_file is None))

if run_btn:
    try:
        with st.spinner("Running audit..."):
            report_bytes = run_audit(adp_file.getvalue(), uzio_file.getvalue())

        st.success("Report generated (3 tabs).")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_name = f"ADP_vs_UZIO_FIT_SIT_Mismatch_Report_{ts}.xlsx"

        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=report_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
