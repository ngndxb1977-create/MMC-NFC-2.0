import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import io
import base64

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(
    page_title="Mitsubishi Financial Matrix Calculator",
    layout="wide"
)

# --------------------------------------------------
# CUSTOM UX/UI STYLING
# --------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://googleapis.com');
    .stApp {
        background-color: #FBF9F6 !important;
        color: #191919 !important;
        font-family: 'Karma', serif !important;
        font-size: 16px !important;
        line-height: 1.6 !important;
    }
    h1, [data-testid="stHeader"] {
        font-family: 'Quicksand', sans-serif !important;
        font-weight: 700 !important;
        font-size: 2.25rem !important;
        color: #191919 !important;
        letter-spacing: -0.02em;
    }
    h2, h3, h4, h5, h6 {
        font-family: 'Amethysta', serif !important;
        font-weight: 400 !important;
        font-size: 1.5rem !important;
        color: #383838 !important;
        margin-top: 1.5rem !important;
    }
    [data-testid="stSidebar"] {
        background-color: #F4F0EA !important;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        font-family: 'Quicksand', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1.2rem !important;
    }
    [data-testid="stMetricValue"] {
        font-family: 'Amethysta', serif !important;
        font-size: 1.6rem !important;
        font-weight: 400 !important;
        color: #191919 !important;
    }
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p {
        font-family: 'Quicksand', sans-serif !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: #555555 !important;
    }
    .requirement-box {
        background-color: #F4F0EA;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #191919;
        margin-top: 1.5rem;
    }
    .requirement-box ul {
        margin-bottom: 1rem;
        padding-left: 1.25rem;
    }
    .disclaimer-text {
        font-size: 0.85rem;
        color: #555555;
        margin-top: 0.5rem;
        line-height: 1.4;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------
def get_clean_model_name(raw_name):
    raw_name_upper = str(raw_name).upper()
    if 'ATTRAGE' in raw_name_upper or 'ATG' in raw_name_upper:
        return 'Attrage'
    elif 'MIRAGE' in raw_name_upper or 'MIG' in raw_name_upper:
        return 'Mirage'
    elif 'ASX' in raw_name_upper:
        return 'ASX'
    elif 'ECLIPSE' in raw_name_upper:
        return 'Eclipse Cross'
    elif 'XPANDER' in raw_name_upper:
        return 'Xpander'
    elif 'OUTLANDER' in raw_name_upper:
        return 'Outlander'
    elif 'MONTERO' in raw_name_upper:
        return 'Montero Sport'
    elif 'DESTINATOR' in raw_name_upper or 'DST' in raw_name_upper or 'DES' in raw_name_upper:
        return 'Destinator'
    return str(raw_name)

def normalize_bracket_string(raw_val):
    if not raw_val or pd.isna(raw_val):
        return ""
    val = str(raw_val).strip().replace(",", "")
    val = re.sub(r'\s+', '', val)
    val = val.replace("to", "-").replace("—", "-").replace("–", "-")
    return val

def get_image_base64(image_path):
    if image_path and os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        ext = os.path.splitext(image_path)[1].lower().replace('.', '')
        mime_type = "image/jpeg" if ext in ["jpg", "jpeg"] else "image/png"
        return f"data:{mime_type};base64,{encoded_string}"
    return None

def safe_float(val, default=0.0):
    try:
        if pd.isna(val) or str(val).strip() == "":
            return default
        cleaned = str(val).replace("AED", "").replace(",", "").replace("%", "").strip()
        return float(cleaned)
    except:
        return default

def find_sheet_name(xls, candidates):
    existing = {s.strip().upper(): s for s in xls.sheet_names}
    for cand in candidates:
        if cand.strip().upper() in existing:
            return existing[cand.strip().upper()]
    return None

def get_col_by_keyword(row, columns, keywords, fallback_idx=None, default=0.0):
    for col in columns:
        col_norm = re.sub(r'\s+', '', str(col)).upper()
        if all(kw.upper() in col_norm for kw in keywords):
            return safe_float(row.get(col), default=default)
    if fallback_idx is not None and len(row) > fallback_idx:
        return safe_float(row.iloc[fallback_idx], default=default)
    return default

# --------------------------------------------------
# MASTER EXCEL EXTRACTION ENGINE (UPDATED SCHEMA)
# --------------------------------------------------
@st.cache_data
def load_master_workbook(file_path):
    data = {
        "price_list": {},
        "bank_rates": {},
        "accessories": {},
        "discounts": {},
        "reservation_fees": {},
        "insurance": {},
        "rmc": {}
    }
    if not os.path.exists(file_path):
        st.error(f"Master workbook not found: {file_path}")
        return data

    xls = pd.ExcelFile(file_path)

    # 1. MASTER PRICE LIST
    mpl_sheet = find_sheet_name(xls, ["Master Price List", "MASTER_PRICE_LIST"])
    if mpl_sheet:
        df_mpl = pd.read_excel(xls, sheet_name=mpl_sheet)
        df_mpl.columns = [str(c).strip() for c in df_mpl.columns]
        cols = list(df_mpl.columns)
        for _, row in df_mpl.iterrows():
            opt_code = str(row.get("OPTION CODE", row.iloc[0] if len(row) > 0 else "")).strip()
            desc = str(row.get("DESCRIPTION", row.iloc[1] if len(row) > 1 else "")).strip()
            var_sap = str(row.get("VARIANT AS IN SAP", row.iloc[2] if len(row) > 2 else "")).strip()
            if not desc or desc.lower() in ["nan", "description"]:
                continue

            raw_year = str(row.get("MODEL YEAR", row.iloc[10] if len(row) > 10 else "")).strip()
            year_match = re.search(r'(20\d{2})', raw_year)
            year_key = year_match.group(1) if year_match else (raw_year if raw_year and raw_year.lower() != "nan" else "2026")

            price_val = safe_float(row.get("PRICE", row.iloc[3] if len(row) > 3 else 0.0))
            bm_disc_a = get_col_by_keyword(row, cols, ["BM", "DISCOUNTA"], fallback_idx=4, default=0.0)
            bm_disc_b = get_col_by_keyword(row, cols, ["BM", "DISCOUNTB"], fallback_idx=5, default=0.0)
            l3_upgraded = get_col_by_keyword(row, cols, ["L3"], fallback_idx=6, default=0.0)
            final_price_val = safe_float(row.get("Final Price", row.iloc[7] if len(row) > 7 else price_val))
            if final_price_val == 0.0 and price_val > 0.0:
                final_price_val = price_val

            if year_key not in data["price_list"]:
                data["price_list"][year_key] = []

            data["price_list"][year_key].append({
                "OptionCode": opt_code,
                "Description": desc,
                "VariantSAP": var_sap,
                "Price": price_val,
                "BM_Discount_A": bm_disc_a,
                "BM_Discount_B": bm_disc_b,
                "L3_Upgraded": l3_upgraded,
                "FinalPrice": final_price_val,
                "ModelYear": year_key,
                "CombinedKey": str(row.get("Combined Key Column", "")).strip()
            })

    # 2. BANK RATES (A: Bank Name, B: Salary Bracket, C: ROI)
    bank_sheet = find_sheet_name(xls, ["Bank_Rates", "Bank Details"])
    if bank_sheet:
        df_bank = pd.read_excel(xls, sheet_name=bank_sheet)
        for _, row in df_bank.iterrows():
            bank_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
            bracket_raw = str(row.iloc[1]).strip() if len(row) > 1 else ""
            roi_raw = row.iloc[2] if len(row) > 2 else None
            if not bank_name or bank_name.lower() in ["nan", "bank name"]:
                continue
            if not bracket_raw or bracket_raw.lower() == "nan":
                continue
            parsed_roi = safe_float(roi_raw, default=-1.0)
            if parsed_roi < 0:
                continue
            if parsed_roi > 1.0:
                parsed_roi = parsed_roi / 100.0
            norm_bracket = normalize_bracket_string(bracket_raw)
            if bank_name not in data["bank_rates"]:
                data["bank_rates"][bank_name] = {}
            data["bank_rates"][bank_name][norm_bracket] = parsed_roi

    # 3. ACCESSORIES (A: List Of Accessories, B: Price)
    acc_sheet = find_sheet_name(xls, ["ACESSORIES", "ACCESSORIES"])
    if acc_sheet:
        df_acc = pd.read_excel(xls, sheet_name=acc_sheet)
        for _, row in df_acc.iterrows():
            acc_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not acc_name or acc_name.lower() in ["nan", "list of accessories"]:
                continue
            acc_price = safe_float(row.iloc[1], default=0.0) if len(row) > 1 else 0.0
            data["accessories"][acc_name] = acc_price

    # 4. DISCOUNT (A: Discount Name, B: Percentage)
    disc_sheet = find_sheet_name(xls, ["Discount", "DISCOUNT"])
    if disc_sheet:
        df_disc = pd.read_excel(xls, sheet_name=disc_sheet)
        for _, row in df_disc.iterrows():
            disc_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not disc_name or disc_name.lower() in ["nan", "discount name", "option code"]:
                continue
