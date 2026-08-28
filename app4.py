import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import datetime
import pytz
import io
import sqlite3
import smtplib
from email.message import EmailMessage

# ==============================================================================
# PAGE CONFIGURATION & TIMEZONE
# ==============================================================================
try:
    st.set_page_config(
        page_title="SAP Inventory & Tonnage Hub",
        page_icon="⏳",
        layout="wide",
        initial_sidebar_state="expanded"
    )
except:
    pass

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

# ==============================================================================
# DATABASE INITIALIZATION (PERMANENT BLOB & RULES STORAGE)
# ==============================================================================
def init_db():
    try:
        conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS saved_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                upload_date TEXT,
                file_name TEXT,
                file_blob BLOB
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS material_shelf_mapping (
                material_code TEXT PRIMARY KEY,
                shelf_days INTEGER
            )
        """)
        conn.commit()
        conn.close()
    except:
        pass

init_db()

def get_saved_shelf_mappings():
    try:
        conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
        df_map = pd.read_sql("SELECT material_code, shelf_days FROM material_shelf_mapping", conn)
        conn.close()
        return dict(zip(df_map['material_code'].astype(str).str.strip().str.lower(), df_map['shelf_days']))
    except:
        return {}

def save_multiple_mappings(m_codes_list, s_days):
    try:
        conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
        cursor = conn.cursor()
        for m_code in m_codes_list:
            cursor.execute(
                "INSERT OR REPLACE INTO material_shelf_mapping (material_code, shelf_days) VALUES (?, ?)",
                (str(m_code).strip().lower(), int(s_days))
            )
        conn.commit()
        conn.close()
        return True
    except:
        return False

# ==============================================================================
# SESSION STATE MANAGEMENT
# ==============================================================================
if "active_df" not in st.session_state:
    st.session_state.active_df = None
if "selected_file_id" not in st.session_state:
    st.session_state.selected_file_id = None
if "calc_theme_choice" not in st.session_state:
    st.session_state.calc_theme_choice = "⚡ Cyber Neon Glass"
if "ng_reset_token" not in st.session_state:
    st.session_state.ng_reset_token = 0

# ==============================================================================
# UI STYLING & HIGH-CONTRAST WHITE TEXT LIVE CLOCK HEADER
# ==============================================================================
st.markdown("""
    <style>
        .stApp {
            background-color: #0f172a;
            color: #f8fafc;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .main-hero {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            padding: 24px;
            border-radius: 16px;
            border: 1px solid #334155;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4);
            margin-bottom: 20px;
        }
        .live-clock-box {
            background: #1e293b;
            color: #ffffff !important;
            padding: 14px 22px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: 700;
            border: 2px solid #38bdf8;
            display: inline-block;
            text-align: center;
            box-shadow: 0 4px 15px rgba(56, 189, 248, 0.3);
            width: 100%;
        }
    </style>
""", unsafe_allow_html=True)

clock_html = """
    <div style="width: 100%;">
        <div class="live-clock-box" id="liveClock" style="color: #ffffff !important;">🕒 Initializing Clock...</div>
    </div>
    <script>
        function updateClock() {
            const now = new Date();
            const options = { timeZone: 'Asia/Kolkata', hour12: true, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
            const clockEl = document.getElementById('liveClock');
            clockEl.innerHTML = '🕒 ' + now.toLocaleString('en-IN', options) + ' IST';
            clockEl.style.color = '#ffffff';
        }
        setInterval(updateClock, 1000);
        updateClock();
    </script>
"""

col_h1, col_h2 = st.columns([2, 1])
with col_h1:
    st.markdown("""
        <div class="main-hero" style="margin-bottom:0px; padding:18px;">
            <h2 style="color: #f8fafc; margin: 0;">⏳ SAP Production, Expiry & Tonnage Hub</h2>
            <p style="color: #94a3b8; margin: 6px 0 0 0; font-size: 13px;">
                Permanent BLOB Storage, ID Reset, Vehicle Tonnage (F2 & BAG), Manual Calculator & Executive Email.
            </p>
        </div>
    """, unsafe_allow_html=True)
with col_h2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.components.v1.html(clock_html, height=75)

st.markdown("<br>", unsafe_allow_html=True)

default_shelf_days = st.number_input(
    "⏱️ Default Global Shelf-Life Period (in Days)", 
    min_value=1, 
    max_value=1095, 
    value=180, 
    step=30, 
    key="global_def_days_input"
)

# ==============================================================================
# CORE PROCESSING FUNCTION WITH FIXED SCOPING
# ==============================================================================
def process_dataframe(df_raw, manual_mfg=None, manual_mat=None):
    try:
        date_keywords = ["production date", "mfg date", "production_date", "mfg_date", "mfg", "production"]
        mfg_col_found = manual_mfg
        if not mfg_col_found:
            for c_name in df_raw.columns:
                c_str = str(c_name).strip().lower()
                if any(kw in c_str for kw in date_keywords):
                    mfg_col_found = c_name
                    break

        mat_col_found = manual_mat
        if not mat_col_found:
            for c_name in df_raw.columns:
                c_low = str(c_name).strip().lower()
                if any(k in c_low for k in ["material", "sku", "item", "code", "product"]):
                    # avoid matching product description
                    if "description" not in c_low:
                        mat_col_found = c_name
                        break
            if not mat_col_found:
                for c_name in df_raw.columns:
                    if "product" in str(c_name).lower():
                        mat_col_found = c_name
                        break

        if mfg_col_found is not None:
            today_dt = pd.Timestamp(get_ist_now().date())
            df_raw['Parsed_Mfg_Date'] = pd.to_datetime(df_raw[mfg_col_found], errors='coerce')
            
            db_rules = get_saved_shelf_mappings()
            
            shelf_days_list = []
            for idx, row in df_raw.iterrows():
                assigned_days = default_shelf_days
                if mat_col_found is not None:
                    m_val = str(row[mat_col_found]).strip().lower()
                    if m_val in db_rules:
                        assigned_days = db_rules[m_val]
                shelf_days_list.append(assigned_days)
                
            df_raw['Assigned_Shelf_Days'] = shelf_days_list
            df_raw['Calculated_Expiry_Date'] = df_raw['Parsed_Mfg_Date'] + pd.to_timedelta(df_raw['Assigned_Shelf_Days'], unit='d')
            df_raw['Remaining_Shelf_Life_Days'] = (df_raw['Calculated_Expiry_Date'] - today_dt).dt.days

            conditions = [
                df_raw['Remaining_Shelf_Life_Days'].isna(),
                df_raw['Remaining_Shelf_Life_Days'] < 0,
                df_raw['Remaining_Shelf_Life_Days'] <= 30
            ]
            choices = ["Unknown Date", "🔴 Expired", "🟡 Critical (<30 Days)"]
            df_raw['Shelf_Life_Status'] = np.select(conditions, choices, default="🟢 Fresh Stock")
            
            def make_remark(val):
                if pd.isna(val):
                    return "Review Date"
                elif val < 0:
                    return "Immediate Action: Stock Expired!"
                elif val <= 30:
                    return "Priority Dispatch / Clearance Required"
                else:
                    return "Stock Condition Healthy"
            
            df_raw['Inventory_Remarks'] = df_raw['Remaining_Shelf_Life_Days'].apply(make_remark)
            
            return df_raw, mat_col_found, mfg_col_found
        else:
            return None, None, None
    except Exception as e_proc:
        st.error(f"❌ Processing error: {str(e_proc)}")
        return None, None, None

# ==============================================================================
# LEFT SIDEBAR: MASTER SUITE, UPLOADER & TABLE NAVIGATION WITH ID RESET
# ==============================================================================
with st.sidebar:
    with st.expander("📦 Logistics Master Suite", expanded=True):
        st.markdown("🔹 Inbound Demand Engine")
        st.markdown("🔹 Route Dispatch Planner")
        st.markdown("🔹 Live ERP Matcher")
        st.markdown("🟢 **SAP Expiry & Tonnage Hub**")

    st.markdown("---")
    st.markdown("### 📂 Upload New SAP Export")
    uploaded_sap_file = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "csv"], key="sidebar_uploader")

    if uploaded_sap_file is not None:
        try:
            file_bytes = uploaded_sap_file.getvalue()
            if uploaded_sap_file.name.endswith('.csv'):
                raw_df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                excel_obj = pd.ExcelFile(io.BytesIO(file_bytes))
                sheet_target = 'Sheet1' if 'Sheet1' in excel_obj.sheet_names else excel_obj.sheet_names[0]
                raw_df = pd.read_excel(excel_obj, sheet_name=sheet_target)

            st.markdown("---")
            st.markdown("**🔍 Column Mapping Assistant:**")
            all_cols = list(raw_df.columns)
            mfg_override = st.selectbox("Select Production/Mfg Date Column", options=["-- Auto Detect --"] + all_cols)
            mat_override = st.selectbox("Select Material Code Column", options=["-- Auto Detect --"] + all_cols)

            man_mfg = None if mfg_override == "-- Auto Detect --" else mfg_override
            man_mat = None if mat_override == "-- Auto Detect --" else mat_override

            processed_df, _, _ = process_dataframe(raw_df, manual_mfg=man_mfg, manual_mat=man_mat)
            if processed_df is not None:
                if st.button("💾 Save & Replace in Permanent DB", type="primary"):
                    conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
                    upload_timestamp = get_ist_now().strftime('%Y-%m-%d %H:%M:%S')
                    cursor = conn.cursor()
                    
                    cursor.execute("DELETE FROM saved_files")
                    try:
                        cursor.execute("DELETE FROM sqlite_sequence WHERE name='saved_files'")
                    except:
                        pass

                    cursor.execute(
                        "INSERT INTO saved_files (upload_date, file_name, file_blob) VALUES (?, ?, ?)",
                        (upload_timestamp, uploaded_sap_file.name, sqlite3.Binary(file_bytes))
                    )
                    conn.commit()
                    conn.close()
                    st.session_state.active_df = processed_df
                    st.success("✅ File saved permanently! Old data replaced & IDs reset.")
                    st.rerun()
            else:
                st.warning("⚠️ Production Date column nahi mila. Please Mapping Assistant se select karein.")
        except Exception as err_file:
            st.error(f"❌ Upload error: {str(err_file)}")

    st.markdown("---")
    st.markdown("### 🗂️ Select Table from Database")
    
    try:
        conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
        saved_records = pd.read_sql("SELECT id, upload_date, file_name FROM saved_files", conn)
        conn.close()
    except:
        saved_records = pd.DataFrame()

    if not saved_records.empty:
        for _, row in saved_records.iterrows():
            btn_label = f"📁 [{row['id']}] {row['file_name']}"
            if st.button(btn_label, key=f"tbl_btn_{row['id']}"):
                st.session_state.selected_file_id = row['id']
                try:
                    conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
                    cursor = conn.cursor()
                    cursor.execute("SELECT file_blob, file_name FROM saved_files WHERE id = ?", (row['id'],))
                    row_data = cursor.fetchone()
                    conn.close()
                    if row_data:
                        blob_data, fname = row_data
                        df_from_db = pd.read_csv(io.BytesIO(blob_data)) if fname.endswith('.csv') else pd.read_excel(io.BytesIO(blob_data), sheet_name=0)
                        processed_df, _, _ = process_dataframe(df_from_db)
                        st.session_state.active_df = processed_df
                        st.rerun()
                except Exception as load_err:
                    st.error(f"Load error: {load_err}")
        
        if st.button("🗑️ Delete Table & Reset IDs", type="secondary"):
            conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM saved_files")
            try:
                cursor.execute("DELETE FROM sqlite_sequence WHERE name='saved_files'")
            except:
                pass
            conn.commit()
            conn.close()
            st.session_state.active_df = None
            st.success("🗑️ Table vanished and database IDs reset!")
            st.rerun()
    else:
        st.info("No saved tables found. Upload a file above.")

# ==============================================================================
# AUTO-LOAD FROM DB IF SESSION IS EMPTY
# ==============================================================================
if st.session_state.active_df is None and not saved_records.empty:
    try:
        latest_id = saved_records.iloc[-1]['id']
        conn = sqlite3.connect("inventory_master_hub.db", check_same_thread=False)
        cursor = conn.cursor()
        cursor.execute("SELECT file_blob, file_name FROM saved_files WHERE id = ?", (latest_id,))
        row_data = cursor.fetchone()
        conn.close()
        if row_data:
            blob_data, fname = row_data
            df_from_db = pd.read_csv(io.BytesIO(blob_data)) if fname.endswith('.csv') else pd.read_excel(io.BytesIO(blob_data), sheet_name=0)
            processed_df, _, _ = process_dataframe(df_from_db)
            st.session_state.active_df = processed_df
    except:
        pass

# ==============================================================================
# MAIN DASHBOARD & ADVANCED FILTER PANEL
# ==============================================================================
if st.session_state.active_df is not None:
    st.markdown("### 📊 Inventory Intelligence Dashboard & Filters")

    working_df = st.session_state.active_df.copy()

    mat_col_target = None
    for c_name in working_df.columns:
        if any(k in str(c_name).lower() for k in ["material", "sku", "item", "code", "product"]):
            if "description" not in str(c_name).lower():
                mat_col_target = c_name
                break

    global_search = st.text_input("🔍 Global Keyword Search (Batch, Plant, Material, Description):", "", key="global_search_input")
    if str(global_search).strip() != "":
        term = str(global_search).strip().lower()
        mask = pd.Series(False, index=working_df.index)
        for c in working_df.columns:
            mask = mask | working_df[c].astype(str).str.lower().str.contains(term, na=False)
        working_df = working_df[mask]

    col_f1, col_f2 = st.columns(2)
    
    with col_f1:
        if mat_col_target is not None:
            unique_mats = sorted(working_df[mat_col_target].dropna().astype(str).unique().tolist())
            selected_materials = st.multiselect("🔍 Multi-Select Material Code(s):", options=unique_mats)
            if selected_materials:
                working_df = working_df[working_df[mat_col_target].astype(str).isin(selected_materials)]

    with col_f2:
        all_headers = list(working_df.columns)
        chosen_header_filter = st.selectbox("📌 Select Header to Filter by Value:", options=["-- Select Header --"] + all_headers)
        if chosen_header_filter != "-- Select Header --":
            unique_vals = sorted(working_df[chosen_header_filter].dropna().astype(str).unique().tolist())
            selected_header_vals = st.multiselect(f"Select value(s) for `{chosen_header_filter}`:", options=unique_vals)
            if selected_header_vals:
                working_df = working_df[working_df[chosen_header_filter].astype(str).isin(selected_header_vals)]

    if "Remaining_Shelf_Life_Days" in working_df.columns:
        max_days_val = int(working_df['Remaining_Shelf_Life_Days'].dropna().max()) if not working_df['Remaining_Shelf_Life_Days'].dropna().empty else 365
        min_days_val = int(working_df['Remaining_Shelf_Life_Days'].dropna().min()) if not working_df['Remaining_Shelf_Life_Days'].dropna().empty else -100
        
        with st.expander("📅 Filter by Expiry Timeline Range (Remaining Days)", expanded=False):
            day_range = st.slider("Select Remaining Shelf-Life Days Range:", min_value=min_days_val, max_value=max(365, max_days_val), value=(min_days_val, max(365, max_days_val)))
            working_df = working_df[(working_df['Remaining_Shelf_Life_Days'] >= day_range[0]) & (working_df['Remaining_Shelf_Life_Days'] <= day_range[1])]

    if mat_col_target is not None:
        with st.expander("🛠️ Update Shelf-Life Rule for Multiple Materials (Permanent Save)"):
            with st.form("rule_update_form_final"):
                mat_list_exp = sorted(working_df[mat_col_target].dropna().astype(str).unique().tolist())
                sel_mat_codes = st.multiselect("Select Material Code(s)", options=mat_list_exp)
                
                shelf_choices = {
                    "30 Days": 30,
                    "60 Days": 60,
                    "90 Days": 90,
                    "180 Days": 180,
                    "365 Days (1 Year)": 365,
                    "548 Days (1.5 Years)": 548,
                    "730 Days (2 Years)": 730
                }
                sel_shelf_label = st.selectbox("Select Shelf Life Period", options=list(shelf_choices.keys()))
                submit_rule_btn = st.form_submit_button("💾 Save Rule for Selected Materials", type="primary")

                if submit_rule_btn and sel_mat_codes:
                    chosen_days = shelf_choices[sel_shelf_label]
                    save_multiple_mappings(sel_mat_codes, chosen_days)
                    
                    full_df = st.session_state.active_df.copy()
                    today_dt = pd.Timestamp(get_ist_now().date())
                    for idx, row in full_df.iterrows():
                        if str(row[mat_col_target]).strip() in sel_mat_codes:
                            full_df.loc[idx, 'Assigned_Shelf_Days'] = chosen_days
                    
                    full_df['Calculated_Expiry_Date'] = full_df['Parsed_Mfg_Date'] + pd.to_timedelta(full_df['Assigned_Shelf_Days'], unit='d')
                    full_df['Remaining_Shelf_Life_Days'] = (full_df['Calculated_Expiry_Date'] - today_dt).dt.days
                    
                    conditions = [full_df['Remaining_Shelf_Life_Days'].isna(), full_df['Remaining_Shelf_Life_Days'] < 0, full_df['Remaining_Shelf_Life_Days'] <= 30]
                    choices = ["Unknown Date", "🔴 Expired", "🟡 Critical (<30 Days)"]
                    full_df['Shelf_Life_Status'] = np.select(conditions, choices, default="🟢 Fresh Stock")
                    
                    st.session_state.active_df = full_df
                    st.success(f"✅ Rules updated successfully for selected materials!")
                    st.rerun()

    if "Shelf_Life_Status" in working_df.columns:
        status_counts = working_df["Shelf_Life_Status"].value_counts()
        m1, m2, m3 = st.columns(3)
        m1.metric("🟢 Fresh Stock", int(status_counts.get("🟢 Fresh Stock", 0)))
        m2.metric("🟡 Critical (<30 Days)", int(status_counts.get("🟡 Critical (<30 Days)", 0)))
        m3.metric("🔴 Expired", int(status_counts.get("🔴 Expired", 0)))

        # ======================================================================
        # VEHICLE TONNAGE CALCULATOR (FROM UPLOADED FILE DATA - F2 & BAG)
        # ======================================================================
        with st.expander("🚚 Vehicle Tonnage Calculator (Billing Type F2 & BAG Slabs)", expanded=True):
            st.markdown("Select Vehicle Number and Billing Date to instantly compute exact Precision vs Standard Tonnage based on your billing records.")
            
            veh_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["vehicle", "truck", "veh"]):
                    veh_col = c
                    break
            
            date_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["billing date", "date", "inv date", "posting"]):
                    date_col = c
                    break
            
            btype_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["billing type", "btype", "type"]):
                    btype_col = c
                    break

            qty_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["invoiced quantity", "billing quantity", "invoice qty", "qty", "quantity"]):
                    qty_col = c
                    break

            unit_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["sale unit", "unit", "uom"]):
                    unit_col = c
                    break

            desc_col = None
            for c in working_df.columns:
                if any(k in str(c).lower() for k in ["product description", "material description", "desc", "item description", "text"]):
                    desc_col = c
                    break

            if veh_col and qty_col:
                calc_df = working_df.copy()
                if btype_col:
                    calc_df = calc_df[calc_df[btype_col].astype(str).str.upper().str.contains("F2", na=False)]
                if unit_col:
                    calc_df = calc_df[calc_df[unit_col].astype(str).str.upper().str.contains("BAG", na=False)]

                unique_vehicles = sorted(calc_df[veh_col].dropna().astype(str).unique().tolist()) if not calc_df.empty else []
                
                if unique_vehicles:
                    sel_vehicle = st.selectbox("🚛 Select Vehicle Number:", options=unique_vehicles, key="calc_veh_select")
                    
                    veh_subset = calc_df[calc_df[veh_col].astype(str) == sel_vehicle]
                    unique_dates = sorted(veh_subset[date_col].dropna().astype(str).unique().tolist()) if date_col and not veh_subset.empty else ["All Dates"]
                    
                    sel_date = st.selectbox("📅 Select Billing Date:", options=unique_dates, key="calc_date_select")
                    
                    if date_col and sel_date != "All Dates":
                        final_veh_rows = veh_subset[veh_subset[date_col].astype(str) == sel_date]
                    else:
                        final_veh_rows = veh_subset

                    b50_total = 0.0
                    b25_total = 0.0
                    
                    for _, r in final_veh_rows.iterrows():
                        q = float(r[qty_col]) if pd.notna(r[qty_col]) else 0.0
                        d_text = str(r[desc_col]).upper() if desc_col and pd.notna(r[desc_col]) else ""
                        
                        if "25" in d_text:
                            b25_total += q
                        else:
                            b50_total += q

                    w1_opt1 = b50_total * 50.120 + b25_total * 25.120
                    mt1 = w1_opt1 / 1000.0

                    w1_opt2 = b50_total * 50.0 + b25_total * 25.0
                    mt2 = w1_opt2 / 1000.0

                    st.markdown(f"**Vehicle Summary for `{sel_vehicle}` (Total Bags: {int(b50_total + b25_total):,}):**")
                    c_res1, c_res2 = st.columns(2)
                    with c_res1:
                        st.metric("🔹 Precision Scale (Opt 1: 50.12 & 25.12)", f"{mt1:,.3f} MT", f"{w1_opt1:,.2f} Kgs")
                    with c_res2:
                        st.metric("🔹 Standard Slabs (Opt 2: 50.0 & 25.0)", f"{mt2:,.3f} MT", f"{w1_opt2:,.2f} Kgs")
                else:
                    st.info("ℹ️ No records found matching Billing Type 'F2' and Unit 'BAG'.")
            else:
                st.warning("⚠️ Vehicle Number or Invoiced Quantity column could not be automatically detected in this file.")

        # ======================================================================
        # MANUAL BAG INPUT CALCULATOR
        # ======================================================================
        with st.expander("⚡ Next-Gen Enterprise Manual Tonnage Calculator", expanded=False):
            CALC_THEMES = {
                "⚡ Cyber Neon Glass": {
                    "bg": "#0b0f19", "card_bg": "rgba(30, 41, 59, 0.75)", "border": "#38bdf8", 
                    "accent": "#38bdf8", "text": "#f8fafc", "subtext": "#94a3b8", "radius": "16px"
                },
                "🌲 Emerald Corporate": {
                    "bg": "#064e3b", "card_bg": "rgba(6, 78, 59, 0.85)", "border": "#34d399", 
                    "accent": "#34d399", "text": "#ecfdf5", "subtext": "#a7f3d0", "radius": "10px"
                },
                "🍇 Executive Burgundy": {
                    "bg": "#500724", "card_bg": "rgba(80, 7, 36, 0.85)", "border": "#f472b6", 
                    "accent": "#f472b6", "text": "#fdf2f8", "subtext": "#fbcfe8", "radius": "20px"
                },
                "💼 Classic Navy Slate": {
                    "bg": "#0f172a", "card_bg": "rgba(15, 23, 42, 0.9)", "border": "#64748b", 
                    "accent": "#60a5fa", "text": "#ffffff", "subtext": "#cbd5e1", "radius": "12px"
                }
            }

            selected_theme_name = st.selectbox("🎨 Select Calculator UI Theme & Layout Style:", list(CALC_THEMES.keys()), key="calc_theme_dropdown")
            st.session_state.calc_theme_choice = selected_theme_name
            th = CALC_THEMES[selected_theme_name]

            st.markdown(f"""
                <style>
                    .nextgen-card {{
                        background: {th['card_bg']};
                        backdrop-filter: blur(12px);
                        padding: 22px;
                        border-radius: {th['radius']};
                        border: 1px solid {th['border']};
                        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.2);
                        margin-bottom: 15px;
                    }}
                    .card-header {{
                        color: {th['accent']};
                        font-size: 17px;
                        font-weight: 700;
                        margin-bottom: 12px;
                    }}
                    .metric-value {{
                        font-size: 26px;
                        font-weight: 800;
                        color: {th['text']};
                        margin: 6px 0;
                    }}
                    .sub-metric {{
                        font-size: 13px;
                        color: {th['subtext']};
                    }}
                </style>
            """, unsafe_allow_html=True)

            k_f1 = f"ng_input_50kg_{st.session_state.ng_reset_token}"
            k_f2 = f"ng_input_25kg_{st.session_state.ng_reset_token}"

            st.markdown(f"<h4 style='color: {th['text']};'>📥 Manual Bag Quantity Inputs</h4>", unsafe_allow_html=True)
            col_inp1, col_inp2 = st.columns(2)

            with col_inp1:
                bags_50 = st.number_input("📦 Category A (50 Kg Bags)", min_value=0, max_value=1000000, step=10, value=0, key=k_f1)
            with col_inp2:
                bags_25 = st.number_input("📦 Category B (25 Kg Bags)", min_value=0, max_value=1000000, step=10, value=0, key=k_f2)

            total_bags_combined = int(bags_50) + int(bags_25)

            w1_m_opt1 = bags_50 * 50.120 + bags_25 * 25.120
            mt_m_opt1 = w1_m_opt1 / 1000.0

            w1_m_opt2 = bags_50 * 50.0 + bags_25 * 25.0
            mt_m_opt2 = w1_m_opt2 / 1000.0

            st.markdown("<br>", unsafe_allow_html=True)
            res_col1, res_col2 = st.columns(2)

            with res_col1:
                st.markdown(f"""
                    <div class="nextgen-card">
                        <div class="card-header">🔹 Option 1 (Precision Scale: 50.12 & 25.12)</div>
                        <div class="sub-metric">Total Volume: <b style="color: {th['text']};">{total_bags_combined:,} Bags</b></div>
                        <hr style="border-color: {th['border']}; margin: 10px 0; opacity: 0.4;">
                        <div class="sub-metric">Combined Net Weight:</div>
                        <div class="metric-value">{w1_m_opt1:,.3f} <span style="font-size: 14px; color: {th['subtext']};">Kgs</span></div>
                        <div style="color: #34d399; font-weight: 700; font-size: 15px; margin-top: 4px;">🚀 {mt_m_opt1:,.3f} Metric Tons (MT)</div>
                    </div>
                """, unsafe_allow_html=True)

            with res_col2:
                st.markdown(f"""
                    <div class="nextgen-card">
                        <div class="card-header">🔹 Option 2 (Standard Slabs: 50.0 & 25.0)</div>
                        <div class="sub-metric">Total Volume: <b style="color: {th['text']};">{total_bags_combined:,} Bags</b></div>
                        <hr style="border-color: {th['border']}; margin: 10px 0; opacity: 0.4;">
                        <div class="sub-metric">Combined Net Weight:</div>
                        <div class="metric-value">{w1_m_opt2:,.2f} <span style="font-size: 14px; color: {th['subtext']};">Kgs</span></div>
                        <div style="color: #34d399; font-weight: 700; font-size: 15px; margin-top: 4px;">🚀 {mt_m_opt2:,.3f} Metric Tons (MT)</div>
                    </div>
                """, unsafe_allow_html=True)

            if st.button("🔄 Reset Manual Calculator State", type="secondary"):
                st.session_state.ng_reset_token += 1
                st.rerun()

        with st.expander("📈 View Expiry Risk & Stock Health Visual Analytics", expanded=False):
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                fig_pie = px.pie(
                    names=status_counts.index, 
                    values=status_counts.values, 
                    title="Stock Health Distribution",
                    hole=0.4,
                    color=status_counts.index,
                    color_discrete_map={"🟢 Fresh Stock": "#22c55e", "🟡 Critical (<30 Days)": "#eab308", "🔴 Expired": "#ef4444", "Unknown Date": "#94a3b8"}
                )
                fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
                st.plotly_chart(fig_pie, use_container_width=True)
            with chart_col2:
                fig_bar = px.bar(
                    x=status_counts.index, 
                    y=status_counts.values, 
                    title="Stock Count by Status Category",
                    labels={'x': 'Status Category', 'y': 'Total Items'},
                    color=status_counts.index,
                    color_discrete_map={"🟢 Fresh Stock": "#22c55e", "🟡 Critical (<30 Days)": "#eab308", "🔴 Expired": "#ef4444", "Unknown Date": "#94a3b8"}
                )
                fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
                st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.dataframe(working_df, use_container_width=True)

    # ==========================================================================
    # ACTION BUTTONS: HTML PRINT VIEW, EXCEL DOWNLOAD & EXECUTIVE EMAIL DISPATCH
    # ==========================================================================
    st.markdown("---")
    st.markdown("### 🚀 Export, Print & Email Options")
    
    col_act1, col_act2, col_act3 = st.columns(3)

    with col_act1:
        html_table_string = working_df.to_html(classes='table table-striped', index=False, border=0)
        print_html_code = f"""
            <html>
                <head>
                    <title>Inventory Report Print View</title>
                    <style>
                        body {{ font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; color: #1e293b; }}
                        h2 {{ color: #0f172a; margin-bottom: 5px; }}
                        p {{ color: #475569; font-size: 13px; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 11px; table-layout: auto; }}
                        th, td {{ border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; word-wrap: break-word; }}
                        th {{ background-color: #1e293b; color: white; font-weight: 600; }}
                        tr:nth-child(even) {{ background-color: #f8fafc; }}
                        @media print {{
                            body {{ padding: 0; }}
                            table {{ page-break-inside: auto; }}
                            tr {{ page-break-inside: avoid; page-break-after: auto; }}
                        }}
                    </style>
                </head>
                <body>
                    <h2>SAP Inventory Expiry & Stock Intelligence Report</h2>
                    <p><b>Generated on:</b> {get_ist_now().strftime('%d-%m-%Y %H:%M:%S IST')}</p>
                    {html_table_string}
                    <script>
                        window.onload = function() {{ window.print(); }}
                    </script>
                </body>
            </html>
        """
        encoded_html = io.BytesIO(print_html_code.encode('utf-8'))
        st.download_button(
            "🖨️ Print / Save Table as PDF",
            data=encoded_html,
            file_name=f"Printable_Inventory_Report_{get_ist_now().strftime('%Y%m%d_%H%M%S')}.html",
            mime="text/html",
            type="secondary",
            help="Downloads a responsive formatted HTML view that opens cleanly in print layout without cutting columns."
        )

    with col_act2:
        excel_buffer = io.BytesIO()
        working_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button(
            "📥 Download Filtered Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"Inventory_Report_{get_ist_now().strftime('%Y-%m-%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    with col_act3:
        with st.expander("✉️ Send Executive Report via Email (KPI Cards, HTML Summary & Attachment)"):
            with st.form("email_dispatch_form"):
                st.markdown("**SMTP Server Settings (Gmail / Corporate Hub):**")
                smtp_host = st.text_input("SMTP Server", "smtp.gmail.com")
                smtp_port = st.number_input("SMTP Port", value=587, step=1)
                sender_email = st.text_input("Sender Email (From)", "your_email@gmail.com")
                sender_pass = st.text_input("Sender App Password", type="password", help="Use Gmail App Password")
                
                st.markdown("---")
                email_to = st.text_input("Recipient Email(s) separated by comma (,):", "recipient@example.com")
                email_sub = st.text_input("Email Subject", "🚨 Executive SAP Inventory Expiry & Critical Stock Summary")
                
                default_mail_body = (
                    "Dear Leadership / Management Team,\n\n"
                    "Please find below the executive summary card and preview table containing critical, expired, "
                    "and fresh stock analytics from SAP.\n\n"
                    "Kindly review the stock status for immediate clearance and dispatch actions.\n\n"
                    "Best Regards,\n"
                    "Supply Chain Management Hub"
                )
                email_body = st.text_area("Email Message", default_mail_body, height=120)
                
                send_email_btn = st.form_submit_button("📨 Send Professional Executive Email", type="primary")

                if send_email_btn:
                    recipients = [e.strip() for e in email_to.split(",") if e.strip()]
                    if not recipients:
                        st.error("❌ Kripya kam se kam ek valid recipient email enter karein.")
                    else:
                        try:
                            msg = EmailMessage()
                            msg['Subject'] = email_sub
                            msg['From'] = sender_email
                            msg['To'] = ", ".join(recipients)
                            
                            total_items = len(working_df)
                            fresh_cnt = int(status_counts.get("🟢 Fresh Stock", 0))
                            crit_cnt = int(status_counts.get("🟡 Critical (<30 Days)", 0))
                            exp_cnt = int(status_counts.get("🔴 Expired", 0))

                            table_html_snippet = working_df.head(50).to_html(index=False, border=1, classes='styled-table')
                            
                            html_content = f"""
                            <html>
                            <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; line-height: 1.6; padding: 15px;">
                                <div style="background: #1e293b; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                                    <h2 style="margin: 0; color: #38bdf8;">📊 Executive Stock Expiry Intelligence Summary</h2>
                                    <p style="margin: 5px 0 0 0; font-size: 13px; color: #94a3b8;">Automated Report | Generated on {get_ist_now().strftime('%d-%m-%Y %H:%M:%S IST')}</p>
                                </div>
                                
                                <p style="font-size: 14px; white-space: pre-wrap;">{email_body}</p>
                                
                                <h3 style="margin-top: 20px; color: #0f172a; border-bottom: 2px solid #cbd5e1; padding-bottom: 5px;">📌 Executive KPI Overview</h3>
                                <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; text-align: center;">
                                    <tr>
                                        <td style="background: #f1f5f9; border: 1px solid #cbd5e1; padding: 12px; border-radius: 6px;">
                                            <span style="font-size: 12px; color: #64748b; font-weight: bold;">TOTAL ITEMS</span><br>
                                            <span style="font-size: 20px; font-weight: bold; color: #0f172a;">{total_items}</span>
                                        </td>
                                        <td style="background: #dcfce7; border: 1px solid #86efac; padding: 12px; border-radius: 6px;">
                                            <span style="font-size: 12px; color: #166534; font-weight: bold;">🟢 FRESH STOCK</span><br>
                                            <span style="font-size: 20px; font-weight: bold; color: #166534;">{fresh_cnt}</span>
                                        </td>
                                        <td style="background: #fef9c3; border: 1px solid #fde047; padding: 12px; border-radius: 6px;">
                                            <span style="font-size: 12px; color: #854d0e; font-weight: bold;">🟡 CRITICAL (<30 DAYS)</span><br>
                                            <span style="font-size: 20px; font-weight: bold; color: #854d0e;">{crit_cnt}</span>
                                        </td>
                                        <td style="background: #fee2e2; border: 1px solid #fca5a5; padding: 12px; border-radius: 6px;">
                                            <span style="font-size: 12px; color: #991b1b; font-weight: bold;">🔴 EXPIRED</span><br>
                                            <span style="font-size: 20px; font-weight: bold; color: #991b1b;">{exp_cnt}</span>
                                        </td>
                                    </tr>
                                </table>

                                <h3 style="color: #0f172a; border-bottom: 2px solid #cbd5e1; padding-bottom: 5px;">📋 Inventory Snapshot Preview (Top 50 Rows)</h3>
                                <style>
                                    .styled-table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
                                    .styled-table th {{ background-color: #1e293b; color: white; padding: 8px; text-align: left; border: 1px solid #cbd5e1; }}
                                    .styled-table td {{ padding: 7px 8px; border: 1px solid #cbd5e1; text-align: left; color: #334155; }}
                                    .styled-table tr:nth-child(even) {{ background-color: #f8fafc; }}
                                </style>
                                {table_html_snippet}
                                
                                <p style="margin-top: 25px; font-size: 12px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px;">
                                    <i>Note: The complete detailed dataset has been attached as an Excel sheet (`.xlsx`) to this email.</i><br>
                                    <b>Supply Chain Automated Intelligence Hub</b>
                                </p>
                            </body>
                            </html>
                            """
                            msg.set_content(email_body)
                            msg.add_alternative(html_content, subtype='html')

                            excel_data = io.BytesIO()
                            working_df.to_excel(excel_data, index=False)
                            excel_data.seek(0)
                            msg.add_attachment(
                                excel_data.getvalue(),
                                maintype='application',
                                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                filename=f"Inventory_Report_{get_ist_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            )

                            server = smtplib.SMTP(smtp_host, int(smtp_port))
                            server.starttls()
                            server.login(sender_email, sender_pass)
                            server.send_message(msg)
                            server.quit()

                            st.success(f"✅ Professional Executive email with KPI summary cards and Excel attachment successfully dispatched to: **{', '.join(recipients)}**!")
                        except Exception as mail_err:
                            st.error(f"❌ Email sending failed. Error details: {str(mail_err)}")
else:
    st.info("ℹ️ Kripya left sidebar se apni SAP stock export file upload karein ya saved table select karein.")
