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
        "rmc_df": None
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
        data["rmc_df"] = pd.read_excel(xls, sheet_name=rmc_sheet, header=None)
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

fn_find_valid_image = lambda filename: next((p for p in [os.path.join("images", filename), os.path.join("Images", filename), filename] if os.path.exists(p)), None)

MASTER_DATA = load_master_workbook(FILE_MASTER)
PRICE_LIST = MASTER_DATA["price_list"]
BANK_RULES = MASTER_DATA["bank_rates"]
ACCESSORIES_MASTER = MASTER_DATA["accessories"]
DISCOUNT_MASTER = MASTER_DATA["discounts"]
RESERVATION_FEES_MASTER = MASTER_DATA["reservation_fees"]
INSURANCE_MASTER = MASTER_DATA["insurance"]
RMC_RAW_DF = MASTER_DATA["rmc_df"]
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

        standard_discounts = ["None", "Esaad", "Fazaa", "Humat Alwatan", "Emirates Platinum", "Permissible-1", "Permissible-2", "Special Discount", "Custom Discount"]
        disc_options_list = ["None"]
        for d in standard_discounts:
            if d != "None" and d not in disc_options_list:
                disc_options_list.append(d)
        for k in DISCOUNT_MASTER.keys():
            if k not in disc_options_list:
                disc_options_list.append(k)

        selected_discount_name = st.selectbox("Applicable Discount:", disc_options_list)
        discount_amount = 0.0

        if selected_discount_name in ["Esaad", "Fazaa", "Humat Alwatan", "Emirates Platinum"]:
            discount_amount = list_vehicle_price * 0.03
            st.info(f"Applied 3% Discount: AED {discount_amount:,.2f}")
        elif selected_discount_name == "Permissible-1":
            discount_amount = 1000.0
            st.info(f"Permissible-1 Fixed Discount: AED {discount_amount:,.2f}")
        elif selected_discount_name == "Permissible-2":
            discount_amount = st.number_input("Permissible-2 Discount Amount (Max AED 2,000):", value=2000.0, min_value=0.0, max_value=2000.0, step=100.0)
        elif selected_discount_name == "Special Discount":
            discount_amount = st.number_input("Special Discount Amount (AED):", value=3000.0, min_value=0.0, step=100.0)
        elif selected_discount_name == "Custom Discount":
            discount_amount = st.number_input("Custom Discount Amount (AED):", value=4000.0, min_value=0.0, step=100.0)
        elif selected_discount_name != "None" and selected_discount_name in DISCOUNT_MASTER:
            val = DISCOUNT_MASTER[selected_discount_name]
            discount_amount = list_vehicle_price * val if val <= 1.0 else val

        discounted_base_price = max(0.0, list_vehicle_price - discount_amount)

        selected_name = get_clean_model_name(selected_description)
        lookup_name = "Xpander Cross" if (selected_name == "Xpander" and str(selected_option_code).strip().upper() == "XC") else selected_name
        resolved_path = None
        if lookup_name in VEHICLE_IMAGES:
            img_file = VEHICLE_IMAGES[lookup_name]
            resolved_path = fn_find_valid_image(img_file)
        if resolved_path:
            caption_name = "Xpander Cross" if lookup_name == "Xpander Cross" else selected_name
            st.image(resolved_path, caption=f"Mitsubishi {caption_name}", use_container_width=True)

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

        down_payment_pct_val = st.slider("Down Payment Percentage (%):", 0, 100, 20)
        down_payment_pct = down_payment_pct_val / 100.0

        st.markdown("---")
        st.subheader("💳 Down Payment Financing Plan")
        finance_dp_option = st.toggle("Finance the Down Payment?", value=False)

        st.markdown("---")
        st.subheader("➕ Custom Accessories Selection")
        acc_selected_price = 0.0
        ceramic_selected_price = 0.0
        foppfgoldpackage_selected_price = 0.0
        warranty_selected_price = 0.0
        rmc_selected_cost = 0.0
        checked_addons_list = []

        if ACCESSORIES_MASTER:
            acc_options_list = list(ACCESSORIES_MASTER.keys())
            selected_accessories = st.multiselect("Select Accessories:", acc_options_list)
            for acc_name in selected_accessories:
                price_val = ACCESSORIES_MASTER[acc_name]
                display_name = acc_name
                if "FOPPF" in acc_name.upper() or "GOLD PACKAGE" in acc_name.upper():
                    display_name = "FO PPF GOLD PACKAGE"
                    foppfgoldpackage_selected_price += price_val
                elif "CERAMIC" in acc_name.upper() and "WINDOW" in acc_name.upper():
                    ceramic_selected_price += price_val
                elif "WARRANTY" in acc_name.upper():
                    warranty_selected_price += price_val
                else:
                    acc_selected_price += price_val
                checked_addons_list.append({"name": display_name, "price": price_val, "vat_taxable": True})

        # ==========================================
        # EXACT MODEL-BASED RMC LOOKUP (Independent of Model Year)
        # ==========================================
        if RMC_RAW_DF is not None and not RMC_RAW_DF.empty:
            df_r = RMC_RAW_DF.copy()
            
            header_row_idx = 0
            for r_idx in range(min(5, len(df_r))):
                row_str = " ".join([str(val).upper() for val in df_r.iloc[r_idx].values])
                if "RMC-" in row_str or "10-40" in row_str:
                    header_row_idx = r_idx
                    break
            
            rmc_header_keys = [str(val).strip() for val in df_r.iloc[header_row_idx].values if pd.notna(val) and ("RMC-" in str(val).upper() or "KMS" in str(val).upper())]
            if not rmc_header_keys and len(df_r.columns) >= 8:
                rmc_header_keys = [str(df_r.iloc[header_row_idx, c]).strip() for c in range(4, min(8, len(df_r.columns)))]
            
            rmc_dropdown_options = ["NO"] + rmc_header_keys
            chosen_rmc = st.selectbox("Routine Maintenance Contract (RMC):", rmc_dropdown_options)
            
            if chosen_rmc != "NO":
                model_target = str(selected_name).strip().upper()
                b16_val = str(chosen_rmc).strip().upper()
                
                rmc_base_val = 0.0
                try:
                    target_col_idx = None
                    for c_idx in range(len(df_r.columns)):
                        cell_val = str(df_r.iloc[header_row_idx, c_idx]).strip().upper()
                        if b16_val in cell_val or cell_val in b16_val:
                            target_col_idx = c_idx
                            break
                    
                    target_row_idx = None
                    for r_idx in range(header_row_idx + 1, len(df_r)):
                        row_model = str(df_r.iloc[r_idx, 0]).strip().upper()
                        if model_target in row_model or row_model in model_target:
                            target_row_idx = r_idx
                            break

                    if target_row_idx is not None and target_col_idx is not None:
                        rmc_base_val = safe_float(df_r.iloc[target_row_idx, target_col_idx])
                except Exception as e:
                    rmc_base_val = 0.0

                rmc_selected_cost += rmc_base_val
                checked_addons_list.append({
                    "name": f"Routine Maintenance Contract ({chosen_rmc})", 
                    "price": rmc_base_val, 
                    "vat_taxable": True
                })

        u19_valuation_base = (base_vehicle_price + acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price + rmc_selected_cost) * 1.05

        code_clean = selected_option_code.strip()
        
        display_vri_price = float(math.ceil(u19_valuation_base * 3.15 * 1.05 / 100.0))
        
        if code_clean in INSURANCE_MASTER:
            ins_rule = INSURANCE_MASTER[code_clean]
            calc_ins = (u19_valuation_base * ins_rule["pct_rate"] + ins_rule["fixed_amount"]) * ins_rule["vat_multiplier"]
            display_ins_price = float(math.ceil(max(calc_ins, ins_rule["min_premium"])))
        else:
            display_ins_price = 0.0

        is_vri_selected = st.checkbox(f"Vehicle Replacement Insurance (VRI) (+AED {display_vri_price:,.2f})", value=False, key="cb_vri_insurance")
        is_insurance_selected = st.checkbox(f"Vehicle Insurance (+AED {display_ins_price:,.2f})", value=False, key="cb_car_insurance")

        vehicle_insurance_cost = display_ins_price if is_insurance_selected else 0.0
        vri_calculated_cost = display_vri_price if is_vri_selected else 0.0

        if is_vri_selected:
            checked_addons_list.append({"name": "Vehicle Replacement Insurance (VRI)", "price": vri_calculated_cost, "vat_taxable": False})
        if is_insurance_selected:
            checked_addons_list.append({"name": "Vehicle Insurance", "price": vehicle_insurance_cost, "vat_taxable": False})

        if code_clean in RESERVATION_FEES_MASTER:
            vehicle_reservation_fee = RESERVATION_FEES_MASTER[code_clean]
        elif selected_name in ['Attrage', 'Mirage'] or any(tag in code_clean.upper() for tag in ['MIG', 'ATG', 'MOP', 'MOG']):
            vehicle_reservation_fee = 500.0
        else:
            vehicle_reservation_fee = 1000.0

        registration_fee = 600.0
        dp_processing_fee = 315.0
        base_vehicle_vat = base_vehicle_price * 0.05
        total_base_vehicle_price = base_vehicle_price + base_vehicle_vat
        excel_addons_total = acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price + vri_calculated_cost + vehicle_insurance_cost + rmc_selected_cost
        total_vat_charges = (base_vehicle_price + acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price + rmc_selected_cost) * 0.05
        full_vehicle_value_including_addons = base_vehicle_price + excel_addons_total + total_vat_charges
        
        calculated_downpayment = float(math.ceil(full_vehicle_value_including_addons * down_payment_pct))
        finance_amount = full_vehicle_value_including_addons - calculated_downpayment
        bank_processing_fee = float(math.ceil(finance_amount * 0.0105))

        st.write("")
        if st.button("Generate Summary", use_container_width=True):
            st.session_state.view_state = "summary"

if st.session_state.view_state == "input":
    st.title("Mitsubishi Financial Dashboard")
    st.info("Configure variables in the sidebar panel. Then click 'Generate Summary' to view the calculation report.")
elif st.session_state.view_state == "summary":
    st.title("📄 Mitsubishi Financial Matrix Calculator")
    st.subheader(f"Unit Selected: {selected_name} — Option {selected_option_code} ({selected_year})")
    st.markdown(
        f"""
        <div class="details-section">
            <div><strong>Financial Institution:</strong> {selected_bank}</div>
            <div><strong>Flat Interest Rate (ROI):</strong> {bank_rate*100:.4f}%</div>
            <div><strong>Down Payment Ratio:</strong> {down_payment_pct_val}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    if customer_name or selected_sales_exec or fixed_fi_name:
        st.markdown("---")
        col_cust1, col_cust2 = st.columns(2)
        with col_cust1:
            if customer_name:
                st.write(f"Customer Name: {customer_name}")
            if customer_email:
                st.write(f"Customer Email: {customer_email}")
            if customer_mobile:
                st.write(f"Customer Mobile: {customer_mobile}")
        with col_cust2:
            if selected_sales_exec:
                st.write(f"Sales Executive: {selected_sales_exec}")
                st.write(f"Executive Email: {exec_email}")
            if exec_mobile:
                st.write(f"Executive Mobile: {exec_mobile}")
            st.write(f"F&I Specialist: {fixed_fi_name}")
            st.write(f"F&I Email: {fixed_fi_email}")

    lookup_name = "Xpander Cross" if (selected_name == "Xpander" and str(selected_option_code).strip().upper() == "XC") else selected_name
    img_b64 = None
    if lookup_name in VEHICLE_IMAGES:
        img_file = VEHICLE_IMAGES[lookup_name]
        resolved_path = fn_find_valid_image(img_file)
        if resolved_path:
            st.image(resolved_path, width=420)
            img_b64 = get_image_base64(resolved_path)

    st.markdown("---")
    st.header("1. Base Vehicle Price & Tax Breakdown")
    base_price_df = pd.DataFrame([{
        "Vehicle Model / Variant": f"{selected_name} ({selected_option_code})",
        "Base Unit Price (Excl. VAT)": f"AED {base_vehicle_price:,.2f}",
        "VAT (5%)": f"AED {base_vehicle_vat:,.2f}",
        "Total Unit Price (Incl. VAT)": f"AED {total_base_vehicle_price:,.2f}"
    }])
    st.table(base_price_df)

    st.subheader("Financial Overview")
    col_s1, col_s2, col_s3 = st.columns(3)
    col_s1.metric("Total Vehicle Value (incl. Addons)", f"AED {full_vehicle_value_including_addons:,.2f}")
    col_s2.metric("Gross Down Payment Req.", f"AED {calculated_downpayment:,.2f}")
    col_s3.metric("Vehicle Finance Amount", f"AED {finance_amount:,.2f}")

    st.markdown("---")
    st.header("2. Loan Installment Breakdowns")
    st.subheader("🟢 Primary Asset Vehicle Financing")
    tenures = list(range(1, 6))
    vehicle_emi_results = []
    for years in tenures:
        months = years * 12
        total_interest = finance_amount * bank_rate * years
        monthly_emi = (finance_amount + total_interest) / months
        vehicle_emi_results.append({
            "Asset Term (Years)": f"{years} Years ({months} Mos)",
            "Flat ROI %": f"{bank_rate*100:.4f}%",
            "Principal Loan Block": f"AED {finance_amount:,.2f}",
            "Total Interest Accrued": f"AED {total_interest:,.2f}",
            "Monthly Vehicle EMI": f"AED {monthly_emi:,.2f}"
        })
    st.table(pd.DataFrame(vehicle_emi_results))

    dp_results = []
    if finance_dp_option:
        st.write("")
        st.subheader("🔵 Down Payment Loan Financing Options")
        dp_financed_base = max(0.0, calculated_downpayment - vehicle_reservation_fee)
        dp_options = [
            {"months": 3, "rate": 0.0000, "label": "3 Months (0.00% ROI)"},
            {"months": 12, "rate": 0.0525, "label": "12 Months (5.25% ROI)"},
            {"months": 24, "rate": 0.0630, "label": "24 Months (6.30% ROI)"}
        ]
        for opt in dp_options:
            total_interest = dp_financed_base * opt["rate"] * (opt["months"] / 12.0)
            monthly_emi = (dp_financed_base + total_interest) / opt["months"]
            dp_results.append({
                "Term": opt["label"],
                "Financed Balance": f"AED {dp_financed_base:,.2f}",
                "Total Interest": f"AED {total_interest:,.2f}",
                "Monthly EMI": f"AED {monthly_emi:,.2f}"
            })
        st.table(pd.DataFrame(dp_results))

    st.markdown("---")
    st.header("3. Accessories Breakdown")
    addons_table_data = []
    if checked_addons_list:
        total_display_addons_price = 0.0
        for addon in checked_addons_list:
            item_price = addon["price"]
            item_vat = (item_price * 0.05) if addon["vat_taxable"] else 0.0
            total_display_addons_price += (item_price + item_vat)
            addons_table_data.append({
                "Selected Accessories / Services": addon["name"],
                "Individual Price (Base)": f"AED {item_price:,.2f}",
                "VAT Amount (5%)": f"AED {item_vat:,.2f}" if addon["vat_taxable"] else "AED 0.00 (VAT Pre-incl.)",
                "Total Cost (incl. VAT)": f"AED {(item_price + item_vat):,.2f}"
            })
        st.table(pd.DataFrame(addons_table_data))
    else:
        st.write("No optional accessories selected.")

    st.markdown("---")
    st.header("4. Out-of-Pocket Cash Outlay Summary")
    if finance_dp_option:
        grand_total_cash_outlay = vehicle_reservation_fee + registration_fee + dp_processing_fee + bank_processing_fee
    else:
        grand_total_cash_outlay = calculated_downpayment + registration_fee + dp_processing_fee + bank_processing_fee

    col_out1, col_out2 = st.columns(2)
    with col_out1:
        if finance_dp_option:
            st.write(f"Showroom Reservation Fee: AED {vehicle_reservation_fee:,.2f} (Paid Upfront)")
            st.write(f"Remaining Down Payment Balance: AED {max(0.0, calculated_downpayment - vehicle_reservation_fee):,.2f} (Financed via Loan Plan)")
        else:
            st.write(f"Full Down Payment Amount: AED {calculated_downpayment:,.2f} (Upfront Out-of-Pocket)")
        st.write(f"Registration Documentation Fee: AED {registration_fee:,.2f}")
    with col_out2:
        st.write(f"DP Processing Fee (DP PF): AED {dp_processing_fee:,.2f}")
        st.write(f"Bank Processing Fee (Bank PF): AED {bank_processing_fee:,.2f}")

    st.markdown(
        f"""
        <div class="requirement-box">
            <strong>🔑 Actual Upfront Cash Required at Showroom Handover:</strong> AED {grand_total_cash_outlay:,.2f}
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")
    st.header("5. Application Requirements & Disclosures")
    st.markdown(
        r"""
        <div class="requirement-box">
        <strong>📋 Required Documentation Checklist:</strong>
        <ul>
            <li>Passport Copy, Digital Visa & Address Page For Indian Passport, Page #44 For Philippines Passport.</li>
            <li>Emirates ID Card Copy Both Sides.</li>
            <li>Labour Card / Free Zone / Employer ID.</li>
            <li>Copy of the UAE Driver's License Both Sides.</li>
            <li>Current Dated Salary Certificate from The Employer.</li>
            <li>Pay Slips For The Last 3 Months - [If Variance In Salary].</li>
            <li>IBAN.</li>
        </ul>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")
    col_opt1, col_opt2, col_opt3 = st.columns(3)
    with col_opt1:
        if st.button("⬅️ Back To Input Tab", use_container_width=True):
            st.session_state.view_state = "input"
            st.rerun()
    with col_opt2:
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            base_price_df.to_excel(writer, sheet_name='Base Price', index=False)
            pd.DataFrame(vehicle_emi_results).to_excel(writer, sheet_name='Vehicle EMI', index=False)
            if dp_results:
                pd.DataFrame(dp_results).to_excel(writer, sheet_name='DP Financing', index=False)
            if addons_table_data:
                pd.DataFrame(addons_table_data).to_excel(writer, sheet_name='Accessories', index=False)
        excel_data_bytes = output_excel.getvalue()
        st.download_button(
            label="📊 Save Excel",
            data=excel_data_bytes,
            file_name=f"Mitsubishi_Financial_Summary_{selected_option_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_opt3:
        logo_b64 = (
            get_image_base64("mitsubishi_logo.png")
            or get_image_base64("mitsubishi.png")
            or get_image_base64("logo.png")
            or get_image_base64("images/mitsubishi_logo.png")
        )
        upfront_label = 'Showroom Reservation Fee (Paid Upfront)' if finance_dp_option else 'Full Down Payment Amount (Upfront Out-of-Pocket)'
        upfront_amount = vehicle_reservation_fee if finance_dp_option else calculated_downpayment

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Mitsubishi Financial Summary - {selected_name}</title>
            <style>
                @import url('https://googleapis.com');
                * {{
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                    color-adjust: exact !important;
                }}
                @page {{
                    size: A4;
                    margin: 12mm 14mm;
                }}
                body {{
                    background-color: #FBF9F6 !important;
                    color: #191919 !important;
                    font-family: 'Karma', serif;
                    font-size: 13.5px;
                    line-height: 1.45;
                    margin: 0;
                    padding: 15px;
                }}
                .header-container, .details-section, .metrics-grid, table, .requirement-box, .metric-card, .section-block {{
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                }}
                .page-break {{
                    page-break-before: always !important;
                    break-before: page !important;
                    padding-top: 4px;
                }}
                .header-container {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    border-bottom: 2px solid #383838;
                    padding-bottom: 10px;
                    margin-bottom: 12px;
                }}
                .logo-img {{
                    max-height: 45px;
                }}
                .vehicle-img {{
                    max-width: 210px;
                    height: auto;
                    display: block;
                    margin: 6px auto;
                }}
                h1 {{
                    font-family: 'Quicksand', sans-serif;
                    font-weight: 700;
                    font-size: 1.7rem;
                    color: #191919;
                    margin: 0;
                }}
                h2 {{
                    font-family: 'Amethysta', serif;
                    font-weight: 400;
                    color: #383838;
                    margin-top: 0.8rem;
                    margin-bottom: 0.5rem;
                    font-size: 1.25rem;
                    border-bottom: 1px solid #E2DDD5;
                    padding-bottom: 4px;
                }}
                h3 {{
                    font-family: 'Amethysta', serif;
                    font-weight: 400;
                    color: #383838;
                    margin-top: 0.7rem;
                    margin-bottom: 0.4rem;
                    font-size: 1.1rem;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 6px;
                    margin-bottom: 12px;
                    background-color: #FFFFFF !important;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
                }}
                th, td {{
                    padding: 8px 10px;
                    text-align: left;
                    border-bottom: 1px solid #EFECE6;
                    font-family: 'Karma', serif;
                    font-size: 0.88rem;
                }}
                th {{
                    background-color: #F4F0EA !important;
                    font-family: 'Quicksand', sans-serif;
                    font-weight: 600;
                    color: #191919;
                }}
                tr.total-row td {{
                    background-color: #F4F0EA !important;
                    font-weight: 700;
                    border-top: 1.5px solid #383838;
                    color: #191919;
                }}
                .metrics-grid {{
                    display: flex;
                    gap: 12px;
                    margin-top: 8px;
                    margin-bottom: 10px;
                }}
                .metric-card {{
                    background-color: #F4F0EA !important;
                    padding: 12px;
                    border-radius: 8px;
                    flex: 1;
                    border-left: 4px solid #191919;
                }}
                .metric-label {{
                    font-family: 'Quicksand', sans-serif;
                    font-size: 0.82rem;
                    font-weight: 600;
                    color: #555555;
                    margin-bottom: 4px;
                }}
                .metric-val {{
                    font-family: 'Amethysta', serif;
                    font-size: 1.3rem;
                    color: #191919;
                }}
                .requirement-box {{
                    background-color: #F4F0EA !important;
                    padding: 1.1rem 1.2rem;
                    border-radius: 8px;
                    border-left: 4px solid #191919;
                    margin-top: 0.8rem;
                    font-size: 0.9rem;
                }}
                .requirement-box ul {{
                    margin: 8px 0 0 0;
                    padding-left: 1.25rem;
                }}
                .requirement-box li {{
                    margin-bottom: 5px;
                }}
                .details-section {{
                    display: flex;
                    justify-content: space-between;
                    background: #F4F0EA !important;
                    padding: 10px 14px;
                    border-radius: 8px;
                    margin-bottom: 10px;
                    font-size: 0.88rem;
                }}
                .action-bar {{
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    background: #191919;
                    padding: 10px 20px;
                    border-radius: 8px;
                    box-shadow: 0 4px 6px rgba(0,0,0,0.2);
                    z-index: 1000;
                }}
                .action-bar button {{
                    background-color: #FBF9F6;
                    color: #191919;
                    border: none;
                    padding: 10px 20px;
                    font-family: 'Quicksand', sans-serif;
                    font-weight: 700;
                    border-radius: 5px;
                    cursor: pointer;
                    font-size: 1rem;
                }}
                @media print {{
                    body {{ padding: 0; background-color: #FBF9F6 !important; }}
                    .no-print, .action-bar {{ display: none !important; }}
                }}
            </style>
        </head>
        <body>
            <div class="action-bar no-print">
                <button onclick="window.print()">🖨️ Save PDF / Print</button>
            </div>

            <!-- ==================== PAGE 1 ==================== -->
            <div class="header-container">
                <div>
                    <h1>Mitsubishi Finance Plan</h1>
                    <p style="margin: 4px 0 0 0;">Unit Selected: {selected_name} — Option {selected_option_code} ({selected_year})</p>
                </div>
                <div>
                    {f'<img src="{logo_b64}" class="logo-img">' if logo_b64 else '<h3>MITSUBISHI MOTORS</h3>'}
                </div>
            </div>

            {f'<div style="text-align: center; margin-bottom: 8px;"><img src="{img_b64}" class="vehicle-img"></div>' if img_b64 else ''}

            <div class="details-section">
                <div><strong>Financial Institution:</strong> {selected_bank}</div>
                <div><strong>Flat Interest Rate (ROI):</strong> {bank_rate*100:.4f}%</div>
                <div><strong>Down Payment Ratio:</strong> {down_payment_pct_val}%</div>
            </div>

            <div class="details-section">
                <div>
                    <strong>Customer Name:</strong> {customer_name if customer_name else "N/A"}<br>
                    <strong>Customer Email:</strong> {customer_email if customer_email else "N/A"}<br>
                    <strong>Customer Mobile:</strong> {customer_mobile if customer_mobile else "N/A"}
                </div>
                <div>
                    <strong>Sales Executive:</strong> {selected_sales_exec if selected_sales_exec else "N/A"}<br>
                    <strong>Executive Email:</strong> {exec_email}<br>
                    <strong>F&I Specialist:</strong> {fixed_fi_name} ({fixed_fi_email})
                </div>
            </div>

            <div class="section-block">
                <h2>1. Base Vehicle Price & Tax Breakdown</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Vehicle Model / Variant</th>
                            <th>Base Unit Price (Excl. VAT)</th>
                            <th>VAT (5%)</th>
                            <th>Total Unit Price (Incl. VAT)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>{selected_name} ({selected_option_code})</td>
                            <td>AED {base_vehicle_price:,.2f}</td>
                            <td>AED {base_vehicle_vat:,.2f}</td>
                            <td>AED {total_base_vehicle_price:,.2f}</td>
                        </tr>
                    </tbody>
                </table>

                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="metric-label">Total Vehicle Value (incl. Addons)</div>
                        <div class="metric-val">AED {full_vehicle_value_including_addons:,.2f}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Gross Down Payment Req.</div>
                        <div class="metric-val">AED {calculated_downpayment:,.2f}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Vehicle Finance Amount</div>
                        <div class="metric-val">AED {finance_amount:,.2f}</div>
                    </div>
                </div>
            </div>

            <!-- ==================== PAGE 2 ==================== -->
            <div class="page-break">
                <div class="section-block">
                    <h2>2. Loan Installment Breakdowns</h2>
                    <h3>Primary Asset Vehicle Financing</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Asset Term (Years)</th>
                                <th>Flat ROI %</th>
                                <th>Principal Loan Block</th>
                                <th>Total Interest Accrued</th>
                                <th>Monthly Vehicle EMI</th>
                            </tr>
                        </thead>
                        <tbody>
                            {"".join([f"<tr><td>{row['Asset Term (Years)']}</td><td>{row['Flat ROI %']}</td><td>{row['Principal Loan Block']}</td><td>{row['Total Interest Accrued']}</td><td>{row['Monthly Vehicle EMI']}</td></tr>" for row in vehicle_emi_results])}
                        </tbody>
                    </table>
                </div>

                {f'''
                <div class="section-block">
                    <h3>Down Payment Loan Financing Options</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Term</th>
                                <th>Financed Balance</th>
                                <th>Total Interest</th>
                                <th>Monthly EMI</th>
                            </tr>
                        </thead>
                        <tbody>
                            {"".join([f"<tr><td>{row['Term']}</td><td>{row['Financed Balance']}</td><td>{row['Total Interest']}</td><td>{row['Monthly EMI']}</td></tr>" for row in dp_results])}
                        </tbody>
                    </table>
                </div>
                ''' if dp_results else ""}

                <div class="section-block">
                    <h2>3. Accessories Breakdown</h2>
                    <table>
                        <thead>
                            <tr>
                                <th>Selected Accessories / Services</th>
                                <th>Individual Price (Base)</th>
                                <th>VAT Amount (5%)</th>
                                <th>Total Cost (incl. VAT)</th>
                            </tr>
                        </thead>
                        <tbody>
                            {"".join([f"<tr><td>{item['Selected Accessories / Services']}</td><td>{item['Individual Price (Base)']}</td><td>{item['VAT Amount (5%)']}</td><td>{item['Total Cost (incl. VAT)']}</td></tr>" for item in addons_table_data]) if addons_table_data else "<tr><td colspan='4'>No optional accessories selected.</td></tr>"}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- ==================== PAGE 3 ==================== -->
            <div class="page-break">
                <div class="section-block">
                    <h2>4. Out-of-Pocket Cash Outlay Summary</h2>
                    <table>
                        <thead>
                            <tr>
                                <th>Fee Description</th>
                                <th>Amount</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td>{upfront_label}</td><td>AED {upfront_amount:,.2f}</td></tr>
                            <tr><td>Registration Documentation Fee</td><td>AED {registration_fee:,.2f}</td></tr>
                            <tr><td>DP Processing Fee (DP PF)</td><td>AED {dp_processing_fee:,.2f}</td></tr>
                            <tr><td>Bank Processing Fee (Bank PF)</td><td>AED {bank_processing_fee:,.2f}</td></tr>
                            <tr class="total-row">
                                <td>Actual Upfront Cash Required at Showroom Handover</td>
                                <td>AED {grand_total_cash_outlay:,.2f}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <div class="section-block">
                    <h2>5. Application Requirements & Disclosures</h2>
                    <div class="requirement-box">
                        <strong>Required Documentation Checklist:</strong>
                        <ul>
                            <li>Passport Copy, Digital Visa & Address Page For Indian Passport, Page #44 For Philippines Passport.</li>
                            <li>Emirates ID Card Copy Both Sides.</li>
                            <li>Labour Card / Free Zone / Employer ID.</li>
                            <li>Copy of the UAE Driver's License Both Sides.</li>
                            <li>Current Dated Salary Certificate from The Employer.</li>
                            <li>Pay Slips For The Last 3 Months - [If Variance In Salary].</li>
                            <li>IBAN.</li>
                        </ul>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """
        st.download_button(
            label="📥 Download HTML For PDF/PRINT",
            data=html_content,
            file_name=f"Mitsubishi_Financial_Summary_{selected_option_code}.html",
            mime="text/html",
            use_container_width=True
        )
