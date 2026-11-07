import streamlit as st
import pandas as pd
import numpy as np
import os
import re
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
    @import url('https://fonts.googleapis.com/css2?family=Karma:wght@400;600&family=Amethysta&family=Quicksand:wght@500;700&display=swap');
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
    raw_name_upper = raw_name.upper()
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
    return raw_name
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
# --------------------------------------------------
# MASTER EXCEL EXTRACTION ENGINE
# --------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
    bank_data = {}
    rmc_data = {}
    if os.path.exists(file_path):
        try:
            df_bank = pd.read_excel(file_path, sheet_name='Bank Details', header=None)
            for _, row in df_bank.iterrows():
                try:
                    bank_name = ""
                    start_col = 1
                    ignore_headers = ["S.NO", "S. NO", "SR NO", "SR. NO", "SL NO", "SL. NO", "SERIAL", "#"]
                    for i in range(4):
                        if i >= len(row):
                            break
                        val = str(row.iloc[i]).strip()
                        if val and val.lower() != "nan" and not val.replace('.', '', 1).isdigit():
                            if val.upper() not in ignore_headers:
                                bank_name = val
                                start_col = i + 1
                                break
                    if not bank_name or "BANK NAME" in bank_name.upper():
                        continue
                    if bank_name not in bank_data:
                        bank_data[bank_name] = {}
                    num_cols = len(row)
                    for col_idx in range(start_col, num_cols):
                        cell_val = row.iloc[col_idx]
                        if pd.isna(cell_val) or str(cell_val).strip() == "":
                            continue
                        cell_str = str(cell_val).strip()
                        is_salary_pattern = (
                            re.search(r'\d+', cell_str) is not None
                            and any(x in cell_str.lower() for x in ['-', 'to', '+', 'above', 'k', 'min'])
                        )
                        if is_salary_pattern:
                            for look_ahead in range(1, 3):
                                if col_idx + look_ahead < num_cols:
                                    next_val = row.iloc[col_idx + look_ahead]
                                    if pd.notna(next_val) and str(next_val).strip() != "":
                                        try:
                                            roi_str = str(next_val).replace("%", "").strip()
                                            parsed_roi = float(roi_str)
                                            if parsed_roi > 1.0:
                                                parsed_roi = parsed_roi / 100.0
                                            norm_sb_val = normalize_bracket_string(cell_val)
                                            if "-" in norm_sb_val:
                                                parts = norm_sb_val.split("-")
                                                display_label = f"{parts[0]}-{parts[1]}"
                                            else:
                                                display_label = norm_sb_val
                                            bank_data[bank_name][display_label] = parsed_roi
                                            break
                                        except:
                                            continue
                except:
                    continue
        except Exception as e:
            st.error(f"Critical error initializing Bank Details sheet: {e}")
        try:
            df_rmc = pd.read_excel(file_path, sheet_name='RMC')
            for _, row in df_rmc.iterrows():
                v_code = str(row['Variant Codes']).strip()
                if not v_code or v_code == "nan":
                    continue
                rmc_data[v_code] = {
                    "RMC-10-40": float(row['RMC-10-40']) if pd.notna(row['RMC-10-40']) else 0.0,
                    "RMC-10-60": float(row['RMC-10-60']) if pd.notna(row['RMC-10-60']) else 0.0,
                    "RMC-10-70": float(row['RMC-10-70']) if pd.notna(row['RMC-10-70']) else 0.0,
                    "RMC-10-100": float(row['RMC-10-100']) if pd.notna(row['RMC-10-100']) else 0.0,
                }
        except:
            pass
    return bank_data, rmc_data
@st.cache_data
def load_accessories_master(file_path):
    acc_dict = {}
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                name = str(row.iloc[0]).strip()
                try:
                    price = float(row.iloc[1])
                except:
                    price = 0.0
                if name and name != "nan":
                    acc_dict[name] = price
        except Exception as e:
            st.error(f"Error loading Accessories List file: {e}")
    return acc_dict
@st.cache_data
def load_sales_specialists_master(file_path):
    specialists_dict = {}
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                name = str(row.iloc[0]).strip()
                if name and name != "nan":
                    specialists_dict[name] = {
                        "email": str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else "",
                        "mobile": str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                    }
        except Exception as e:
            st.error(f"Error loading Sales_Specialists file: {e}")
    return specialists_dict
@st.cache_data
def load_all_vehicle_data(vehicle_file_path):
    catalog = {"2025": {}, "2026": {}}
    if not os.path.exists(vehicle_file_path):
        return catalog
    xls = pd.ExcelFile(vehicle_file_path)
    for sheet in xls.sheet_names:
        if sheet.strip() in ['Structure', 'Bank Details', 'RMC'] or 'COMBINED' in sheet.upper():
            continue
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            raw_name = str(df.iloc[4, 2]).strip()
            code = str(df.iloc[4, 3]).strip()
            year_string = str(df.iloc[4, 4]).strip()
            if not raw_name or raw_name == "nan":
                raw_name = sheet
            if not code or code == "nan":
                code = sheet
            year_key = "2026" if "2026" in year_string or "26" in sheet else "2025"
            base_price = float(df.iloc[6, 1])
            interest_rate = float(df.iloc[18, 3])
            registration_fee = float(df.iloc[11, 7]) if pd.notna(df.iloc[11, 7]) else 600.0
            processing_fee_dp = float(df.iloc[12, 7]) if pd.notna(df.iloc[12, 7]) else 315.0
            model_name = get_clean_model_name(raw_name)
            sheet_upper = sheet.upper()
            if model_name in ['Attrage', 'Mirage'] or 'MIG' in sheet_upper or 'ATG' in sheet_upper or 'MOP' in sheet_upper or 'MOG' in sheet_upper:
                reservation_fee = 500.0
            else:
                reservation_fee = 1000.0
            accessories = {}
            row_labels = {
                10: "Custom Accessories", 11: "Ceramic Gold Window Tint", 12: "FO PPF Gold Package",
                13: "Extended Warranty", 14: "VRI", 15: "Vehicle Insurance", 16: "RMC"
            }
            for row_idx, default_label in row_labels.items():
                cell_label = str(df.iloc[row_idx - 1, 1]).strip()
                status = str(df.iloc[row_idx - 1, 2]).strip().upper()
                label = cell_label if (cell_label and cell_label != "nan") else default_label
                is_checked = (status == "YES")
                try:
                    price_val = float(df.iloc[row_idx - 1, 3])
                except:
                    price_val = 0.0
                accessories[label] = {
                    "price_raw": price_val,
                    "default_checked": is_checked,
                    "type_tag": "VRI" if row_idx == 14 else (
                        "INSURANCE" if row_idx == 15 else (
                            "RMC" if row_idx == 16 else "STANDARD"
                        )
                    )
                }
            if model_name not in catalog[year_key]:
                catalog[year_key][model_name] = {}
            catalog[year_key][model_name][code] = {
                "base_price": base_price,
                "interest_rate": interest_rate,
                "accessories": accessories,
                "registration_fee": registration_fee,
                "processing_fee_dp": processing_fee_dp,
                "reservation_fee": reservation_fee
            }
        except:
            continue
    return catalog
# --------------------------------------------------
# PRICE_LIST LOADER (MASTER VEHICLE SOURCE)
# --------------------------------------------------
@st.cache_data
def load_price_list(file_path):
    price_catalog = {"2025": [], "2026": []}
    if not os.path.exists(file_path):
        st.error(f"PRICE_LIST file not found: {file_path}")
        return price_catalog
    try:
        sheet_2026 = "PRICE_LIST_MY_2026"
        sheet_2025 = "Sheet2"
        df_2026 = pd.read_excel(file_path, sheet_name=sheet_2026)
        for _, row in df_2026.iterrows():
            price_catalog["2026"].append({
                "Description": str(row.get("Description", "")).strip(),
                "VariantSAP": str(row.get("VARIANT AS IN SAP", "")).strip(),
                "OptionCode": str(row.get("OTPION CODE", "")).strip(),
                "Price": float(row.get("PRICE", 0))
            })
        df_2025 = pd.read_excel(file_path, sheet_name=sheet_2025)
        for _, row in df_2025.iterrows():
            price_catalog["2025"].append({
                "Description": str(row.get("Description", "")).strip(),
                "VariantSAP": str(row.get("VARIANT AS IN SAP", "")).strip(),
                "OptionCode": str(row.get("OTPION CODE", "")).strip(),
                "Price": float(row.get("PRICE", 0))
            })
    except Exception as e:
        st.error(f"Error loading PRICE_LIST file: {e}")
    return price_catalog
# --------------------------------------------------
# FILE CONSTANTS
# --------------------------------------------------
FILE_VEHICLES = "NFC New VRI Project (2) (2).xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"
FILE_ACCESSORIES = "Accessories_List.xlsx"
FILE_SALES_SPECIALISTS = "Sales_Specialists.xlsx"
FILE_PRICE_LIST = "PRICE_LIST.xlsx"
# --------------------------------------------------
# VEHICLE IMAGES MAP
# --------------------------------------------------
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
# --------------------------------------------------
# LOAD DATASETS
# --------------------------------------------------
VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)
ACCESSORIES_MASTER = load_accessories_master(FILE_ACCESSORIES)
SALES_SPECIALISTS_MASTER = load_sales_specialists_master(FILE_SALES_SPECIALISTS)
PRICE_LIST = load_price_list(FILE_PRICE_LIST)
# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------
if "view_state" not in st.session_state:
    st.session_state.view_state = "input"
# --------------------------------------------------
# SIDEBAR: CUSTOMER, SALES, CONFIG
# --------------------------------------------------
if not VEHICLE_CATALOG["2025"] and not VEHICLE_CATALOG["2026"]:
    st.error(f"Could not load vehicle datasets from '{FILE_VEHICLES}'.")
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
        # --------------------------------------------------
        # CONFIGURATION CONSOLE (PRICE_LIST AS MASTER)
        # --------------------------------------------------
        st.header("🚗 Configuration Console")
        selected_year = st.selectbox("Model Year:", ["2025", "2026"])
        year_rows = PRICE_LIST[selected_year]
        if not year_rows:
            st.error(f"No PRICE_LIST data found for {selected_year}.")
            st.stop()
        desc_options = sorted(list(set([row["Description"] for row in year_rows if row["Description"]]))
        )
        selected_description = st.selectbox("Description:", desc_options)
        filtered_rows = [row for row in year_rows if row["Description"] == selected_description]
        if not filtered_rows:
            st.error("Selected Description not found in PRICE_LIST. Check mapping.")
            st.stop()
        final_row = filtered_rows[0]
        selected_variant_sap = final_row["VariantSAP"]
        selected_option_code = final_row["OptionCode"]
        base_vehicle_price = final_row["Price"]
        st.write(f"**Variant Code (SAP):** {selected_variant_sap}")
        st.write(f"**Option Code:** {selected_option_code}")
        st.write(f"**Base Price (from PRICE_LIST):** AED {base_vehicle_price:,.2f}")
        selected_name = get_clean_model_name(selected_description)
        model_catalog = VEHICLE_CATALOG[selected_year].get(selected_name, {})
        if not model_catalog:
            st.error(f"No VEHICLE_CATALOG data found for model '{selected_name}' in year {selected_year}.")
            st.stop()
        first_variant_key = list(model_catalog.keys())[0]
        v_data = model_catalog[first_variant_key]
        lookup_name = "Xpander Cross" if (selected_name == "Xpander" and str(selected_option_code).strip().upper() == "XC") else selected_name
        resolved_path = None
        if lookup_name in VEHICLE_IMAGES:
            img_file = VEHICLE_IMAGES[lookup_name]
            resolved_path = find_valid_image_path(img_file)
        if resolved_path:
            caption_name = "Xpander Cross" if lookup_name == "Xpander Cross" else selected_name
            st.image(resolved_path, caption=f"Mitsubishi {caption_name}", use_container_width=True)
        else:
            st.sidebar.error(f"⚠️ Missing: Put '{VEHICLE_IMAGES.get(lookup_name, 'image.png')}' in your 'images' folder.")
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
            fetched_roi = v_data["interest_rate"]
        bank_rate = st.number_input("Flat Interest Rate (ROI):", value=fetched_roi, format="%.4f", step=0.0001)
        st.markdown("---")
        st.subheader("⚙️ Calculations Adjuster")
        base_vehicle_price = st.number_input(
            "Base Vehicle Price (AED):",
            value=base_vehicle_price,
            step=500.0
        )
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
        else:
            st.warning("Accessories_List file not found or empty.")
            
        if RMC_RULES and selected_option_code in RMC_RULES:
            rmc_keys = list(RMC_RULES[selected_option_code].keys())
        else:
            rmc_keys = ["RMC-10-40", "RMC-10-60", "RMC-10-70", "RMC-10-100"]
            
        rmc_dropdown_options = ["None"] + rmc_keys
        chosen_rmc = st.selectbox("Routine Maintenance Contract (RMC):", rmc_dropdown_options)
        
        if chosen_rmc != "None":
            if RMC_RULES and selected_option_code in RMC_RULES and chosen_rmc in RMC_RULES[selected_option_code]:
                rmc_selected_cost = RMC_RULES[selected_option_code][chosen_rmc]
            else:
                for name, info in v_data["accessories"].items():
                    if info["type_tag"] == "RMC":
                        rmc_selected_cost = info["price_raw"]
                        break
            checked_addons_list.append({
                "name": f"Routine Maintenance Contract ({chosen_rmc})",
                "price": rmc_selected_cost / 1.05,
                "vat_taxable": True
            })
                
        u19_valuation_base = (
            base_vehicle_price
            + acc_selected_price
            + ceramic_selected_price
            + foppfgoldpackage_selected_price
            + warranty_selected_price
            + (rmc_selected_cost / 1.05)
        ) * 1.05
        is_vri_selected = False
        is_insurance_selected = False
        display_vri_price = u19_valuation_base * 3.15 * 1.05 / 100
        code_clean = selected_option_code.strip()
        name_clean = str(selected_name)
        if code_clean.startswith(('PR', 'HLP')) or code_clean in [
            "OTF03", "OTP03", "OTF06", "OTP06", "OTP06-01",
            "OTF08", "OTP08", "G03", "G05", "G06", "G08",
            "G09", "G10", "G12", "P03", "P06", "P12", "G31"
        ]:
            display_ins_price = (u19_valuation_base * 0.03 + 510) * 1.05
        elif code_clean.startswith(('H', 'P', 'MOP', 'MOG')) and any(
            num in code_clean for num in ["57", "59", "61", "62", "64"]
        ):
            display_ins_price = (u19_valuation_base * 0.0275 + 510) * 1.05
        elif code_clean.startswith('EH') and any(num in code_clean for num in ["40", "41", "43"]):
            display_ins_price = (u19_valuation_base * 0.03 + 450) * 1.05
        else:
            display_ins_price = 3690.0 if "Xpander" in name_clean or "Destinator" in name_clean else 3625.0
        for name, info in v_data["accessories"].items():
            if info["type_tag"] == "VRI":
                is_vri_selected = st.checkbox(
                    f"Vehicle Replacement Insurance (VRI) (+AED {display_vri_price:,.2f})",
                    value=info["default_checked"],
                    key="cb_vri_insurance"
                )
            elif info["type_tag"] == "INSURANCE":
                is_insurance_selected = st.checkbox(
                    f"Vehicle Insurance (+AED {display_ins_price:,.2f})",
                    value=info["default_checked"],
                    key="cb_car_insurance"
                )
        vehicle_insurance_cost = display_ins_price if is_insurance_selected else 0.0
        vri_calculated_cost = display_vri_price if is_vri_selected else 0.0
        if is_vri_selected:
            checked_addons_list.append({
                "name": "Vehicle Replacement Insurance (VRI)",
                "price": vri_calculated_cost,
                "vat_taxable": False
            })
        if is_insurance_selected:
            checked_addons_list.append({
                "name": "Vehicle Insurance",
                "price": vehicle_insurance_cost,
                "vat_taxable": False
            })
        base_vehicle_vat = base_vehicle_price * 0.05
        total_base_vehicle_price = base_vehicle_price + base_vehicle_vat
        excel_addons_total = (
            acc_selected_price
            + ceramic_selected_price
            + foppfgoldpackage_selected_price
            + warranty_selected_price
            + vri_calculated_cost
            + vehicle_insurance_cost
            + rmc_selected_cost
        )
        total_vat_charges = (
            base_vehicle_price
            + acc_selected_price
            + ceramic_selected_price
            + foppfgoldpackage_selected_price
            + warranty_selected_price
        ) * 0.05
        full_vehicle_value_including_addons = (
            base_vehicle_price
            + excel_addons_total
            + total_vat_charges
        )
        calculated_downpayment = full_vehicle_value_including_addons * down_payment_pct
        finance_amount = full_vehicle_value_including_addons - calculated_downpayment
        dp_processing_fee = v_data["processing_fee_dp"]
        bank_processing_fee = finance_amount * 0.0105
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Generate Summary", use_container_width=True):
            st.session_state.view_state = "summary"
# --------------------------------------------------
# MAIN WORKSPACE RENDERING
# --------------------------------------------------
if st.session_state.view_state == "input":
    st.title("Mitsubishi Financial Dashboard")
    st.info(
        "Configure variables in the sidebar panel. Then click **'Generate Summary'** "
        "to view the calculation report."
    )
elif st.session_state.view_state == "summary":
    st.title("📄 Mitsubishi Financial Matrix Calculator")
    st.subheader(f"Unit Selected: {selected_name} — Option {selected_option_code} ({selected_year})")
    
    st.markdown(
        f"""
        <div style="background-color: #F4F0EA; padding: 15px 20px; border-radius: 8px; border-left: 4px solid #191919; margin-bottom: 20px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 15px;">
            <div>
                <span style="font-family: 'Quicksand', sans-serif; font-size: 0.85rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Financial Institution</span>
                <span style="font-family: 'Amethysta', serif; font-size: 1.2rem; color: #191919;">{selected_bank}</span>
            </div>
            <div>
                <span style="font-family: 'Quicksand', sans-serif; font-size: 0.85rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Flat Interest Rate (ROI)</span>
                <span style="font-family: 'Amethysta', serif; font-size: 1.2rem; color: #191919;">{bank_rate*100:.4f}%</span>
            </div>
            <div>
                <span style="font-family: 'Quicksand', sans-serif; font-size: 0.85rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Down Payment Ratio</span>
                <span style="font-family: 'Amethysta', serif; font-size: 1.2rem; color: #191919;">{down_payment_pct_val}%</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    if customer_name or selected_sales_exec or fixed_fi_name:
        st.markdown("---")
        col_cust1, col_cust2 = st.columns(2)
        with col_cust1:
            if customer_name:
                st.write(f"**Customer Name:** {customer_name}")
            if customer_email:
                st.write(f"**Customer Email:** {customer_email}")
            if customer_mobile:
                st.write(f"**Customer Mobile:** {customer_mobile}")
        with col_cust2:
            if selected_sales_exec:
                st.write(f"**Sales Executive:** {selected_sales_exec}")
            st.write(f"**Executive Email:** {exec_email}")
            if exec_mobile:
                st.write(f"**Executive Mobile:** {exec_mobile}")
            st.write(f"**F&I Specialist:** {fixed_fi_name}")
            st.write(f"**F&I Email:** {fixed_fi_email}")
    lookup_name = "Xpander Cross" if (selected_name == "Xpander" and str(selected_option_code).strip().upper() == "XC") else selected_name
    img_b64 = None
    if lookup_name in VEHICLE_IMAGES:
        img_file = VEHICLE_IMAGES[lookup_name]
        resolved_path = find_valid_image_path(img_file)
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
    tenures = [1, 2, 3, 4, 5]
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
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("🔵 Down Payment Loan Financing Options")
        vehicle_reservation_fee = v_data["reservation_fee"]
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
        st.write("*No optional accessories selected.*")
    st.markdown("---")
    st.header("4. Out-of-Pocket Cash Outlay Summary")
    registration_fee = v_data["registration_fee"]
    if finance_dp_option:
        vehicle_reservation_fee = v_data["reservation_fee"]
        grand_total_cash_outlay = vehicle_reservation_fee + registration_fee + dp_processing_fee + bank_processing_fee
    else:
        grand_total_cash_outlay = calculated_downpayment + registration_fee + dp_processing_fee + bank_processing_fee
    col_out1, col_out2 = st.columns(2)
    with col_out1:
        if finance_dp_option:
            st.write(
                f"**Showroom Reservation Fee:** AED {v_data['reservation_fee']:,.2f} (Paid Upfront)"
            )
            st.write(
                f"**Remaining Down Payment Balance:** "
                f"AED {max(0.0, calculated_downpayment - v_data['reservation_fee']):,.2f} "
                f"(Financed via Loan Plan)"
            )
        else:
            st.write(
                f"**Full Down Payment Amount:** AED {calculated_downpayment:,.2f} "
                f"(Upfront Out-of-Pocket)"
            )
        st.write(f"**Registration Documentation Fee:** AED {registration_fee:,.2f}")
    with col_out2:
        st.write(f"**DP Processing Fee (DP PF):** AED {dp_processing_fee:,.2f}")
        st.write(f"**Bank Processing Fee (Bank PF):** AED {bank_processing_fee:,.2f}")
    st.markdown(
        f"""
        <div style="background-color: #F4F0EA; padding: 1.25rem 1.5rem; border-radius: 8px; border-left: 4px solid #191919; margin-top: 1.5rem;">
            <span style="font-family: 'Quicksand', sans-serif; font-weight: 700; font-size: 0.9rem; color: #555555; display: block; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.25rem;">
                🔑 Actual Upfront Cash Required at Showroom Handover
            </span>
            <span style="font-family: 'Amethysta', serif; font-size: 1.8rem; color: #191919;">
                <span style="font-size: 1.2rem;">AED</span> {grand_total_cash_outlay:,.2f}
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("---")
    st.header("5. Application Requirements & Disclosures")
    st.markdown(
        r"""
        <div style="background-color: #F4F0EA; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #191919; margin-top: 1.5rem;">
            <strong style="font-family: sans-serif; font-size: 1.1rem; color: #191919; display: block; margin-bottom: 0.75rem;">📋 Required Documentation Checklist:</strong>
            <ul style="margin-bottom: 1rem; padding-left: 1.25rem;">
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
    # --------------------------------------------------
    # SUMMARY OPTIONS BAR (BACK, EXCEL, HTML)
    # --------------------------------------------------
    st.markdown("---")
    col_opt1, col_opt2, col_opt3 = st.columns(3)
    with col_opt1:
        if st.button("⬅️ Back To Input Tab", use_container_width=True):
            st.session_state.view_state = "input"
            st.rerun()
    with col_opt2:
        import io
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
        logo_b64 = get_image_base64("mitsubishi_logo.png") or get_image_base64("mitsubishi.png") or get_image_base64("logo.png") or get_image_base64("images/mitsubishi_logo.png")
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Mitsubishi Financial Summary - {selected_name}</title>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Karma:wght@400;600&family=Amethysta&family=Quicksand:wght@500;700&display=swap');
                
                * {{
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                    color-adjust: exact !important;
                }}
                
                @page {{
                    size: A4;
                    margin: 15mm;
                }}

                body {{
                    background-color: #FBF9F6 !important;
                    color: #191919 !important;
                    font-family: 'Karma', serif;
                    font-size: 14px;
                    line-height: 1.5;
                    margin: 0;
                    padding: 20px;
                }}
                
                .header-container, .details-section, .metrics-grid, table, .requirement-box, .metric-card {{
                    page-break-inside: avoid !important;
                    break-inside: avoid !important;
                }}

                /* Page 2 Break Enforcement */
                .page-break-page2 {{
                    page-break-before: always !important;
                    break-before: page !important;
                }}

                /* Page 3 Break Enforcement */
                .page-break-page3 {{
                    page-break-before: always !important;
                    break-before: page !important;
                }}

                .header-container {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    border-bottom: 2px solid #383838;
                    padding-bottom: 15px;
                    margin-bottom: 20px;
                }}
                .logo-img {{
                    max-height: 50px;
                }}
                h1 {{
                    font-family: 'Quicksand', sans-serif;
                    font-weight: 700;
                    font-size: 1.8rem;
                    color: #191919;
                    margin: 0;
                }}
                h2, h3 {{
                    font-family: 'Amethysta', serif;
                    font-weight: 400;
                    color: #383838;
                    margin-top: 1.2rem;
                    font-size: 1.3rem;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 10px;
                    margin-bottom: 15px;
                    background-color: #FFFFFF !important;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                }}
                th, td {{
                    padding: 10px 12px;
                    text-align: left;
                    border-bottom: 1px solid #EFECE6;
                    font-family: 'Karma', serif;
                    font-size: 0.9rem;
                }}
                th {{
                    background-color: #F4F0EA !important;
                    font-family: 'Quicksand', sans-serif;
                    font-weight: 600;
                    color: #191919;
                }}
                .metrics-grid {{
                    display: flex;
                    gap: 15px;
                    margin-bottom: 15px;
                }}
                .metric-card {{
                    background-color: #F4F0EA !important;
                    padding: 15px;
                    border-radius: 8px;
                    flex: 1;
                    border-left: 4px solid #191919;
                }}
                .metric-label {{
                    font-family: 'Quicksand', sans-serif;
                    font-size: 0.85rem;
                    font-weight: 600;
                    color: #555555;
                    margin-bottom: 5px;
                }}
                .metric-val {{
                    font-family: 'Amethysta', serif;
                    font-size: 1.4rem;
                    color: #191919;
                }}
                .vehicle-container {{
                    text-align: center;
                    margin: 15px 0;
                }}
                .vehicle-img {{
                    max-width: 320px;
                    height: auto;
                    border-radius: 8px;
                }}
                .requirement-box {{
                    background-color: #F4F0EA !important;
                    padding: 1.2rem;
                    border-radius: 8px;
                    border-left: 4px solid #191919;
                    margin-top: 1rem;
                    font-size: 0.9rem;
                }}
                .details-section {{
                    display: flex;
                    justify-content: space-between;
                    background: #F4F0EA !important;
                    padding: 12px 15px;
                    border-radius: 8px;
                    margin-bottom: 15px;
                    font-size: 0.9rem;
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

            <div class="header-container">
                <div>
                    <h1>Mitsubishi Financial Report</h1>
                    <h3>Unit Selected: {selected_name} — Option {selected_option_code} ({selected_year})</h3>
                </div>
                <div>
                    {f'<img class="logo-img" src="{logo_b64}"/>' if logo_b64 else '<b>MITSUBISHI MOTORS</b>'}
                </div>
            </div>

            <div style="background-color: #F4F0EA !important; padding: 12px 15px; border-radius: 8px; border-left: 4px solid #191919; margin-bottom: 15px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                <div>
                    <span style="font-family: 'Quicksand', sans-serif; font-size: 0.75rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Financial Institution</span>
                    <span style="font-family: 'Amethysta', serif; font-size: 1.1rem; color: #191919;">{selected_bank}</span>
                </div>
                <div>
                    <span style="font-family: 'Quicksand', sans-serif; font-size: 0.75rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Flat Interest Rate (ROI)</span>
                    <span style="font-family: 'Amethysta', serif; font-size: 1.1rem; color: #191919;">{bank_rate*100:.4f}%</span>
                </div>
                <div>
                    <span style="font-family: 'Quicksand', sans-serif; font-size: 0.75rem; font-weight: 600; color: #555555; display: block; text-transform: uppercase;">Down Payment Ratio</span>
                    <span style="font-family: 'Amethysta', serif; font-size: 1.1rem; color: #191919;">{down_payment_pct_val}%</span>
                </div>
            </div>

            <div class="details-section">
                <div>
                    <p style="margin: 4px 0;"><strong>Customer Name:</strong> {customer_name if customer_name else "N/A"}</p>
                    <p style="margin: 4px 0;"><strong>Customer Email:</strong> {customer_email if customer_email else "N/A"}</p>
                    <p style="margin: 4px 0;"><strong>Customer Mobile:</strong> {customer_mobile if customer_mobile else "N/A"}</p>
                </div>
                <div>
                    <p style="margin: 4px 0;"><strong>Sales Executive:</strong> {selected_sales_exec if selected_sales_exec else "N/A"}</p>
                    <p style="margin: 4px 0;"><strong>Executive Email:</strong> {exec_email}</p>
                    <p style="margin: 4px 0;"><strong>F&I Specialist:</strong> {fixed_fi_name}</p>
                </div>
            </div>

            {f'<div class="vehicle-container"><img class="vehicle-img" src="{img_b64}"/></div>' if img_b64 else ''}

            <h2>1. Base Vehicle Price & Tax Breakdown</h2>
            <table>
                <tr>
                    <th>Vehicle Model / Variant</th>
                    <th>Base Unit Price (Excl. VAT)</th>
                    <th>VAT (5%)</th>
                    <th>Total Unit Price (Incl. VAT)</th>
                </tr>
                <tr>
                    <td>{selected_name} ({selected_option_code})</td>
                    <td>AED {base_vehicle_price:,.2f}</td>
                    <td>AED {base_vehicle_vat:,.2f}</td>
                    <td>AED {total_base_vehicle_price:,.2f}</td>
                </tr>
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

            <!-- PAGE 2: LOAN INSTALLMENT BREAKDOWNS -->
            <div class="page-break-page2"></div>
            <h2>2. Loan Installment Breakdowns</h2>
            <h3>Primary Asset Vehicle Financing</h3>
            <table>
                <tr>
                    <th>Asset Term (Years)</th>
                    <th>Flat ROI %</th>
                    <th>Principal Loan Block</th>
                    <th>Total Interest Accrued</th>
                    <th>Monthly Vehicle EMI</th>
                </tr>
                {"".join([f"<tr><td>{row['Asset Term (Years)']}</td><td>{row['Flat ROI %']}</td><td>AED {finance_amount:,.2f}</td><td>AED {finance_amount * bank_rate * (int(row['Asset Term (Years)'].split()[0])):,.2f}</td><td>AED {((finance_amount + (finance_amount * bank_rate * (int(row['Asset Term (Years)'].split()[0])))) / (int(row['Asset Term (Years)'].split()[0]) * 12)):,.2f}</td></tr>" for row in vehicle_emi_results])}
            </table>
            
            {"<h3>Down Payment Loan Financing Options</h3><table><tr><th>Term</th><th>Financed Balance</th><th>Total Interest</th><th>Monthly EMI</th></tr>" + "".join([f"<tr><td>{row['Term']}</td><td>AED {dp_financed_base:,.2f}</td><td>AED {float(row['Total Interest'].replace('AED ','').replace(',','')):,.2f}</td><td>AED {float(row['Monthly EMI'].replace('AED ','').replace(',','')):,.2f}</td></tr>" for row in dp_results]) + "</table>" if dp_results else ""}

            <h2>3. Accessories Breakdown</h2>
            <table>
                <tr>
                    <th>Selected Accessories / Services</th>
                    <th>Individual Price (Base)</th>
                    <th>VAT Amount (5%)</th>
                    <th>Total Cost (incl. VAT)</th>
                </tr>
                {"".join([f"<tr><td>{item['Selected Accessories / Services']}</td><td>AED {float(item['Individual Price (Base)'].replace('AED ','').replace(',','')):,.2f}</td><td>" + (f"AED {float(item['VAT Amount (5%)'].replace('AED ','').replace(',','')):,.2f}" if "VAT Pre-incl" not in item['VAT Amount (5%)'] else "AED 0.00 (VAT Pre-incl.)") + f"</td><td>AED {float(item['Total Cost (incl. VAT)'].replace('AED ','').replace(',','')):,.2f}</td></tr>" for item in addons_table_data]) if addons_table_data else "<tr><td colspan='4'>No optional accessories selected.</td></tr>"}
            </table>

            <!-- PAGE 3: OUT-OF-POCKET CASH OUTLAY SUMMARY -->
            <div class="page-break-page3"></div>
            <h2>4. Out-of-Pocket Cash Outlay Summary</h2>
            <table>
                <tr>
                    <th>Fee Description</th>
                    <th>Amount</th>
                </tr>
                <tr>
                    <td>{'Showroom Reservation Fee (Paid Upfront)' if finance_dp_option else 'Full Down Payment Amount (Upfront Out-of-Pocket)'}</td>
                    <td>AED {v_data["reservation_fee"]:,.2f}</td>
                </tr>
                <tr>
                    <td>Registration Documentation Fee</td>
                    <td>AED {registration_fee:,.2f}</td>
                </tr>
                <tr>
                    <td>DP Processing Fee (DP PF)</td>
                    <td>AED {dp_processing_fee:,.2f}</td>
                </tr>
                <tr>
                    <td>Bank Processing Fee (Bank PF)</td>
                    <td>AED {bank_processing_fee:,.2f}</td>
                </tr>
            </table>
            
            <div class="metric-card" style="margin-top: 15px;">
                <div class="metric-label">Actual Upfront Cash Required at Showroom Handover</div>
                <div class="metric-val">AED {grand_total_cash_outlay:,.2f}</div>
            </div>

            <h2>5. Application Requirements & Disclosures</h2>
            <div class="requirement-box">
                <strong style="display: block; margin-bottom: 5px;">Required Documentation Checklist:</strong>
                <ul style="margin: 0; padding-left: 20px;">
                    <li>Passport Copy, Digital Visa & Address Page For Indian Passport, Page #44 For Philippines Passport.</li>
                    <li>Emirates ID Card Copy Both Sides.</li>
                    <li>Labour Card / Free Zone / Employer ID.</li>
                    <li>Copy of the UAE Driver's License Both Sides.</li>
                    <li>Current Dated Salary Certificate from The Employer.</li>
                    <li>Pay Slips For The Last 3 Months - [If Variance In Salary].</li>
                    <li>IBAN.</li>
                </ul>
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
