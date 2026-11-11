import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import io
import base64
import math

st.set_page_config(
    page_title="Mitsubishi Financial Matrix Calculator",
    layout="wide"
)

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

    acc_sheet = find_sheet_name(xls, ["ACESSORIES", "ACCESSORIES"])
    if acc_sheet:
        df_acc = pd.read_excel(xls, sheet_name=acc_sheet)
        for _, row in df_acc.iterrows():
            acc_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not acc_name or acc_name.lower() in ["nan", "list of accessories"]:
                continue
            acc_price = safe_float(row.iloc[1], default=0.0) if len(row) > 1 else 0.0
            data["accessories"][acc_name] = acc_price

    disc_sheet = find_sheet_name(xls, ["Discount", "DISCOUNT"])
    if disc_sheet:
        df_disc = pd.read_excel(xls, sheet_name=disc_sheet)
        for _, row in df_disc.iterrows():
            disc_name = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not disc_name or disc_name.lower() in ["nan", "discount name", "option code"]:
                continue
            pct_val = safe_float(row.iloc[1], default=0.0) if len(row) > 1 else 0.0
            if pct_val > 1.0:
                pct_val = pct_val / 100.0
            data["discounts"][disc_name] = pct_val

    res_sheet = find_sheet_name(xls, ["Reservation Fees", "RESERVATION FEES"])
    if res_sheet:
        df_res = pd.read_excel(xls, sheet_name=res_sheet)
        df_res.columns = [str(c).strip() for c in df_res.columns]
        for _, row in df_res.iterrows():
            opt_code = str(row.get("OPTION CODE", row.iloc[0] if len(row) > 0 else "")).strip()
            if not opt_code or opt_code.lower() in ["nan", "option code"]:
                continue
            if "Reservation Fees" in df_res.columns:
                res_fee = safe_float(row.get("Reservation Fees"), default=1000.0)
            else:
                res_fee = safe_float(row.iloc[10], default=1000.0) if len(row) > 10 else 1000.0
            data["reservation_fees"][opt_code] = res_fee

    ins_sheet = find_sheet_name(xls, ["Insurance", "INSURANCE"])
    if ins_sheet:
        df_ins = pd.read_excel(xls, sheet_name=ins_sheet)
        df_ins.columns = [str(c).strip() for c in df_ins.columns]
        for _, row in df_ins.iterrows():
            opt_code = str(row.get("OPTION CODE", row.iloc[0] if len(row) > 0 else "")).strip()
            if not opt_code or opt_code.lower() in ["nan", "option code"]:
                continue
            pct_rate = safe_float(row.get("Percentage Rate", row.iloc[10] if len(row) > 10 else 0.0))
            if pct_rate > 1.0:
                pct_rate = pct_rate / 100.0
            fixed_amt = safe_float(row.get("Fixed Amount", row.iloc[11] if len(row) > 11 else 0.0))
            vat_mult = safe_float(row.get("Vat Multiplier", row.iloc[12] if len(row) > 12 else 1.05), default=1.05)
            if vat_mult == 0.0:
                vat_mult = 1.05
            min_prem = safe_float(row.get("Minimum Premium", row.iloc[13] if len(row) > 13 else 0.0))
            data["insurance"][opt_code] = {
                "pct_rate": pct_rate,
                "fixed_amount": fixed_amt,
                "vat_multiplier": vat_mult,
                "min_premium": min_prem
            }

    rmc_sheet = find_sheet_name(xls, ["RMC"])
    if rmc_sheet:
        df_rmc = pd.read_excel(xls, sheet_name=rmc_sheet)
        df_rmc.columns = [str(c).strip() for c in df_rmc.columns]
        for _, row in df_rmc.iterrows():
            v_code = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not v_code or v_code.lower() in ["nan", "variant codes", "option code"]:
                continue
            
            rmc_entry = {}
            for col in df_rmc.columns:
                col_upper = col.upper()
                if "RMC" in col_upper:
                    rmc_entry[col.strip()] = safe_float(row.get(col, 0.0))
            
            if not rmc_entry:
                rmc_entry = {
                    "RMC-10-40": safe_float(row.get("RMC-10-40", 0.0)),
                    "RMC-10-60": safe_float(row.get("RMC-10-60", 0.0)),
                    "RMC-10-70": safe_float(row.get("RMC-10-70", 0.0)),
                    "RMC-10-100": safe_float(row.get("RMC-10-100", 0.0)),
                }
            data["rmc"][v_code.upper()] = rmc_entry
    return data

@st.cache_data
def load_sales_specialists_master(file_path):
    specialists_dict = {}
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                name = str(row.iloc[0]).strip() if len(row) > 0 else ""
                if name and name.lower() not in ["nan", "full name"]:
                    specialists_dict[name] = {
                        "email": str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else "",
                        "mobile": str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                    }
        except Exception as e:
            st.error(f"Error loading Sales_Specialists file: {e}")
    return specialists_dict

FILE_MASTER = "NFC New VRI Project-6.xlsx"
FILE_SALES_SPECIALISTS = "Sales_Specialists.xlsx"

VEHICLE_IMAGES = {
    "Attrage": "attrage.png.png",
    "Mirage": "mirage.png.png",
    "ASX": "asx.png.png",
    "Eclipse Cross": "eclipse_cross.png.png",
    "Xpander": "xpander.png.png",
    "Xpander Cross": "xpander_cross.png.png",
    "Outlander": "outlander.png.png",
    "Montero Sport": "montero.png.png",
    "Destinator": "destinator.png.png"
}

def find_valid_image_path(filename):
    possible_paths = [
        os.path.join("images", filename),
        os.path.join("Images", filename),
        os.path.join("images", filename.replace(".png.png", ".PNG.PNG")),
        os.path.join("Images", filename.replace(".png.png", ".PNG.PNG")),
        filename
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

MASTER_DATA = load_master_workbook(FILE_MASTER)
PRICE_LIST = MASTER_DATA["price_list"]
BANK_RULES = MASTER_DATA["bank_rates"]
ACCESSORIES_MASTER = MASTER_DATA["accessories"]
DISCOUNT_MASTER = MASTER_DATA["discounts"]
RESERVATION_FEES_MASTER = MASTER_DATA["reservation_fees"]
INSURANCE_MASTER = MASTER_DATA["insurance"]
RMC_RULES = MASTER_DATA["rmc"]
SALES_SPECIALISTS_MASTER = load_sales_specialists_master(FILE_SALES_SPECIALISTS)

if "view_state" not in st.session_state:
    st.session_state.view_state = "input"

if not PRICE_LIST:
    st.error(f"Could not load Master Price List from '{FILE_MASTER}'.")
else:
    with st.sidebar:
        st.header("👤 Customer & Sales Details")
        customer_name = st.text_input("Customer Name:")
        customer_email = st.text_input("Email:")
        customer_mobile = st.text_input("Mobile Number:")
        st.markdown("---")

        st.subheader("👨‍💼 Sales Executive")
        if SALES_SPECIALISTS_MASTER:
            sales_exec_names = sorted(list(SALES_SPECIALISTS_MASTER.keys()))
            selected_sales_exec = st.selectbox("Select Sales Executive:", sales_exec_names)
            exec_email = SALES_SPECIALISTS_MASTER[selected_sales_exec]["email"]
            exec_mobile = SALES_SPECIALISTS_MASTER[selected_sales_exec]["mobile"]
        else:
            selected_sales_exec = st.text_input("Sales Executive Name:", value="")
            exec_email = ""
            exec_mobile = ""
        st.text_input("Executive Email:", value=exec_email)
        st.text_input("Executive Mobile:", value=exec_mobile)
        st.markdown("---")

        st.subheader("🏦 Finance & Insurance Specialist")
        fixed_fi_name = "Naveen Gopinath"
        fixed_fi_email = "naveen@habtoormotors.com"
        st.text_input("FI Specialist Name:", value=fixed_fi_name, disabled=True, key="fi_name_display")
        st.text_input("FI Specialist Email:", value=fixed_fi_email, disabled=True, key="fi_email_display")
        st.markdown("---")

        st.header("🚗 Configuration Console")
        available_years = sorted(list(PRICE_LIST.keys()))
        selected_year = st.selectbox("Model Year:", available_years)
        year_rows = PRICE_LIST[selected_year]

        if not year_rows:
            st.error(f"No Master Price List data found for {selected_year}.")
            st.stop()

        desc_options = sorted(list(set([row["Description"] for row in year_rows if row["Description"]])))
        selected_description = st.selectbox("Description:", desc_options)
        filtered_rows = [row for row in year_rows if row["Description"] == selected_description]

        if not filtered_rows:
            st.error("Selected Description not found in Master Price List.")
            st.stop()

        final_row = filtered_rows[0]
        selected_variant_sap = final_row["VariantSAP"]
        selected_option_code = final_row["OptionCode"]
        list_vehicle_price = final_row["Price"]
        bm_disc_a_limit = final_row["BM_Discount_A"]
        bm_disc_b_limit = final_row["BM_Discount_B"]
        l3_upgraded_limit = final_row["L3_Upgraded"]

        st.write(f"**Variant Code (SAP):** {selected_variant_sap}")
        st.write(f"**Option Code:** {selected_option_code}")
        st.write(f"**List Price:** AED {list_vehicle_price:,.2f}")

        selected_discount_name = "None"
        discount_pct = 0.0
        if DISCOUNT_MASTER:
            disc_options = ["None"] + list(DISCOUNT_MASTER.keys())
            selected_discount_name = st.selectbox("Applicable Discount:", disc_options)
            if selected_discount_name != "None":
                discount_pct = DISCOUNT_MASTER[selected_discount_name]

        discounted_base_price = list_vehicle_price * (1.0 - discount_pct)

        selected_name = get_clean_model_name(selected_description)
        lookup_name = "Xpander Cross" if (selected_name == "Xpander" and str(selected_option_code).strip().upper() == "XC") else selected_name
        resolved_path = None
        if lookup_name in VEHICLE_IMAGES:
            img_file = VEHICLE_IMAGES[lookup_name]
            resolved_path = find_valid_image_path(img_file)
        if resolved_path:
            caption_name = "Xpander Cross" if lookup_name == "Xpander Cross" else selected_name
            st.image(resolved_path, caption=f"Mitsubishi {caption_name}", use_container_width=True)
        else:
            st.sidebar.warning(f"⚠️ Image not found: '{VEHICLE_IMAGES.get(lookup_name, 'image.png')}'")

        st.markdown("---")
        st.subheader("🏦 Financial Provider Rates")
        if BANK_RULES:
            bank_options = sorted(list(BANK_RULES.keys()))
            selected_bank = st.selectbox("Select Institution:", bank_options)
            bracket_options = sorted(list(BANK_RULES[selected_bank].keys()))
            selected_bracket = st.selectbox("Income Bracket Selection:", bracket_options)
            fetched_roi = BANK_RULES[selected_bank][selected_bracket]
        else:
            selected_bank = "Sheet Benchmark"
            fetched_roi = 0.0249
        bank_rate = st.number_input("Flat Interest Rate (ROI):", value=fetched_roi, format="%.4f", step=0.0001)

        st.markdown("---")
        st.subheader("⚙️ Calculations Adjuster")
        base_vehicle_price = st.number_input("Base Vehicle Price (AED):", value=float(discounted_base_price), step=500.0)

        applied_discount_amt = max(0.0, list_vehicle_price - base_vehicle_price)
        if applied_discount_amt > 0:
            max_bm_limit = max(bm_disc_a_limit, bm_disc_b_limit)
            max_l3_limit = max(max_bm_limit, l3_upgraded_limit)
            if max_l3_limit > 0 and applied_discount_amt > max_l3_limit:
                st.error(f"🚨 Discount of AED {applied_discount_amt:,.2f} exceeds max L3 limit (AED {max
