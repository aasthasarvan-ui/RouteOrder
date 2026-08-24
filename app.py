import streamlit as st
import pandas as pd
import openpyxl
import datetime
import pytz
import io
import re
import zipfile
import sqlite3
import smtplib
import json
import urllib.parse
from email.message import EmailMessage
from fpdf import FPDF
import streamlit.components.v1 as components

# ==========================================
# 1. PAGE CONFIGURATION & THEME STYLING
# ==========================================
st.set_page_config(
    page_title="Enterprise Sales Order & Dispatch Hub", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

THEMES = {
    "💼 Classic Enterprise Navy": {
        "icon": "💼", "bg": "#f4f6f9", "text": "#1f2937", "card_bg": "#ffffff", "border": "#cbd5e1",
        "btn_bg": "#1e3a8a", "btn_hover": "#1d4ed8", "primary": "#2563eb", "input_bg": "#ffffff", "input_text": "#1f2937"
    },
    "🌙 Modern Dark ERP": {
        "icon": "🌙", "bg": "#0b0f19", "text": "#f3f4f6", "card_bg": "#1f2937", "border": "#374151",
        "btn_bg": "#374151", "btn_hover": "#4b5563", "primary": "#3b82f6", "input_bg": "#111827", "input_text": "#f3f4f6"
    },
    "📊 Corporate Slate": {
        "icon": "📊", "bg": "#eef2f5", "text": "#0f172a", "card_bg": "#ffffff", "border": "#94a3b8",
        "btn_bg": "#475569", "btn_hover": "#334155", "primary": "#0284c7", "input_bg": "#ffffff", "input_text": "#0f172a"
    },
    "☀️ Clean Light Minimal": {
        "icon": "☀️", "bg": "#ffffff", "text": "#111827", "card_bg": "#f9fafb", "border": "#d1d5db",
        "btn_bg": "#0f172a", "btn_hover": "#1e293b", "primary": "#10b981", "input_bg": "#ffffff", "input_text": "#111827"
    },
    "⚡ Cyber Blue": {
        "icon": "⚡", "bg": "#f0fdfa", "text": "#042f2e", "card_bg": "#ccfbf1", "border": "#5eead4",
        "btn_bg": "#0d9488", "btn_hover": "#0f766e", "primary": "#14b8a6", "input_bg": "#ffffff", "input_text": "#042f2e"
    },
    "🌲 Emerald Corporate": {
        "icon": "🌲", "bg": "#f0fdf4", "text": "#14532d", "card_bg": "#dcfce7", "border": "#86efac",
        "btn_bg": "#16a34a", "btn_hover": "#15803d", "primary": "#22c55e", "input_bg": "#ffffff", "input_text": "#14532d"
    },
    "🍇 Executive Burgundy": {
        "icon": "🍇", "bg": "#fdf2f8", "text": "#500724", "card_bg": "#fce7f3", "border": "#f472b6",
        "btn_bg": "#db2777", "btn_hover": "#be185d", "primary": "#ec4899", "input_bg": "#ffffff", "input_text": "#500724"
    },
    "🪙 Titanium Charcoal": {
        "icon": "🪙", "bg": "#18181b", "text": "#fafafa", "card_bg": "#27272a", "border": "#52525b",
        "btn_bg": "#52525b", "btn_hover": "#71717a", "primary": "#e4e4e7", "input_bg": "#09090b", "input_text": "#fafafa"
    }
}

IST = pytz.timezone('Asia/Kolkata')

def get_ist_now():
    return datetime.datetime.now(IST)

# ==========================================
# 2. DATABASE INTEGRITY & INITIALIZATION
# ==========================================
def verify_core_integrity():
    try:
        conn = sqlite3.connect("sales_history.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        existing_tables = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        required_tables = [
            'history_logs', 'unique_routes_master', 'output_files_ledger', 
            'unmapped_missing_dr_ledger', 'input_output_traceability', 
            'discrepancy_audit_ledger', 'dispatch_planning_ledger', 'inventory_stock_table'
        ]
        for t_name in required_tables:
            if t_name not in existing_tables:
                return False, f"Missing critical database table: {t_name}"
        return True, "All Core Integrity Checkpoints Passed Successfully!"
    except Exception as e:
        return False, str(e)

def init_db():
    conn = sqlite3.connect("sales_history.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS history_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, files_count INTEGER, total_qty REAL, status TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS unique_routes_master (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no, dr_code))")
    cursor.execute("CREATE TABLE IF NOT EXISTS output_files_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, file_type TEXT, file_data BLOB, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no))")
    cursor.execute("CREATE TABLE IF NOT EXISTS input_output_traceability (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, input_file_name TEXT, input_file_blob BLOB, total_input_qty REAL, generated_output_file TEXT, output_type TEXT, version_no INTEGER, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS discrepancy_audit_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, file_name TEXT, agency_no TEXT, dr_code TEXT, fg_code TEXT, input_qty REAL, generated_qty REAL, difference REAL, logged_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS dispatch_planning_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, dispatch_date TEXT, route_no TEXT, vehicle_no TEXT, driver_mobile TEXT, agency_no TEXT, fg_code TEXT, demand_qty REAL, dispatched_qty REAL, pending_qty REAL, status TEXT, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS inventory_stock_table (id INTEGER PRIMARY KEY AUTOINCREMENT, fg_code TEXT UNIQUE, available_qty REAL, updated_at TEXT)")
    
    for tbl in ['unique_routes_master', 'dispatch_planning_ledger', 'inventory_stock_table', 'discrepancy_audit_ledger']:
        try:
            cursor.execute(f"PRAGMA table_info({tbl})")
            cols = [col[1] for col in cursor.fetchall()]
            if tbl == 'unique_routes_master' and "file_name" not in cols:
                cursor.execute("ALTER TABLE unique_routes_master ADD COLUMN file_name TEXT")
            if tbl == 'dispatch_planning_ledger' and "fg_code" not in cols:
                cursor.execute("ALTER TABLE dispatch_planning_ledger ADD COLUMN fg_code TEXT")
        except Exception:
            pass
        
    conn.commit()
    conn.close()

init_db()

is_healthy, health_msg = verify_core_integrity()
if not is_healthy:
    st.error(f"❌ **System Integrity Warning:** {health_msg}")
    st.stop()

# ==========================================
# 3. SESSION STATE DEFAULTS
# ==========================================
DEFAULTS = {
    "fg_code": "FG500014",
    "col_map": "36:FG500014AJ\n37:FG500014AK",
    "agency_override": "101:36:FG500014N01\n101:37:FG500014N02",
    "route": "22",
    "email_user": st.secrets.get("email", {}).get("sender_email", ""),
    "email_pass": st.secrets.get("email", {}).get("app_password", ""),
    "recipient": st.secrets.get("email", {}).get("recipient_email", ""),
    "whatsapp": "",
    "selected_theme": "💼 Classic Enterprise Navy",
    "vehicle_capacity": 160.0,
    "vehicle_no": "PB-10-AB-1234",
    "driver_mobile": "9876543210",
    "processed_files": [],
    "comparison_summary": [],
    "skipped_rows_log": [],
    "anomaly_logs": [],
    "unmapped_current_batch": [],
    "kpi_data": {"input_qty": 0, "gen_qty": 0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val

t = THEMES[st.session_state.selected_theme]

st.markdown(
    f"""
    <style>
        #GithubIcon {{ visibility: hidden !important; display: none !important; }}
        .stAppHeader {{ background-color: transparent !important; }}
        header[data-testid="stHeader"] {{ display: none !important; }}
        
        .stApp {{
            background-color: {t['bg']};
            color: {t['text']};
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{
            color: {t['text']} !important;
        }}
        
        input, textarea, select {{
            background-color: {t['input_bg']} !important;
            color: {t['input_text']} !important;
            border: 1px solid {t['border']} !important;
        }}
        
        .stButton>button {{
            width: 100%;
            height: 38px;
            background-color: {t['btn_bg']} !important;
            color: #ffffff !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            border-radius: 4px;
            border: 1px solid {t['border']};
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s ease;
        }}
        .stButton>button p {{ color: #ffffff !important; }}
        .stButton>button:hover {{ background-color: {t['btn_hover']} !important; color: #ffffff !important; }}
        
        button[kind="primary"] {{ background-color: {t['primary']} !important; color: #ffffff !important; }}
        
        div[data-testid="stExpander"] {{
            background-color: {t['card_bg']};
            border: 1px solid {t['border']};
            border-radius: 4px;
        }}
        div[data-testid="stDataFrame"] {{
            border: 1px solid {t['border']};
            border-radius: 4px;
            background-color: {t['card_bg']};
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==========================================
# 4. CONTROL PANEL & STOCK UPLOAD CONFIGURATION
# ==========================================
with st.expander("⚙️ Control Panel, Vehicle Capacity, Stock Inventory & Settings", expanded=True):
    st.subheader("🎨 Theme Engine & Vehicle Dispatch Config")
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        def on_theme_change():
            st.session_state.selected_theme = st.session_state.theme_selectbox
        st.selectbox("Select Interface Theme", list(THEMES.keys()), key="theme_selectbox", index=list(THEMES.keys()).index(st.session_state.selected_theme), on_change=on_theme_change)
    with col_t2:
        st.session_state.vehicle_capacity = st.number_input("Vehicle Capacity (Bags/Cases e.g., 160)", min_value=10.0, max_value=1000.0, value=st.session_state.vehicle_capacity, step=5.0)

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.session_state.vehicle_no = st.text_input("Assigned Vehicle Number (Shared across Route agencies)", value=st.session_state.vehicle_no)
    with col_v2:
        st.session_state.driver_mobile = st.text_input("Driver / Transporter Mobile No", value=st.session_state.driver_mobile)

    st.markdown("---")
    st.subheader("📦 Daily Stock Inventory Upload (For Shortfall / Pending Calculation)")
    stock_upload_file = st.file_uploader("Upload Available Stock Excel/CSV (Columns: fg_code, available_qty)", type=["csv", "xlsx"], key="stock_up")
    if stock_upload_file:
        try:
            if stock_upload_file.name.endswith('.csv'):
                df_stk = pd.read_csv(stock_upload_file)
            else:
                df_stk = pd.read_excel(stock_upload_file)
            
            if 'fg_code' in df_stk.columns and 'available_qty' in df_stk.columns:
                conn_stk = sqlite3.connect("sales_history.db")
                cur_stk = conn_stk.cursor()
                for _, row in df_stk.iterrows():
                    cur_stk.execute("INSERT OR REPLACE INTO inventory_stock_table (fg_code, available_qty, updated_at) VALUES (?, ?, ?)", (str(row['fg_code']), float(row['available_qty']), get_ist_now().strftime("%Y-%m-%d %H:%M:%S")))
                conn_stk.commit()
                conn_stk.close()
                st.success("✅ Stock Inventory successfully uploaded and updated!")
            else:
                st.error("❌ Stock file must contain 'fg_code' and 'available_qty' columns.")
        except Exception as ex:
            st.error(f"Error loading stock file: {str(ex)}")

    st.markdown("---")
    col_set1, col_set2, col_set3 = st.columns(3)
    with col_set1:
        st.subheader("Default Fallback FG Code")
        st.session_state.fg_code = st.text_input("FG Code Input", value=st.session_state.fg_code, label_visibility="collapsed")
        c1, c2 = st.columns(2)
        if c1.button("Clear FG"): st.session_state.fg_code = ""; st.rerun()
        if c2.button("Restore FG"): st.session_state.fg_code = DEFAULTS["fg_code"]; st.rerun()
        
        st.subheader("Default Route Fallback")
        st.session_state.route = st.text_input("Route Input", value=st.session_state.route, label_visibility="collapsed")
        c1, c2 = st.columns(2)
        if c1.button("Clear Route"): st.session_state.route = ""; st.rerun()
        if c2.button("Restore Route"): st.session_state.route = DEFAULTS["route"]; st.rerun()

    with col_set2:
        st.subheader("Direct Column Index Mapping")
        st.session_state.col_map = st.text_area("Col Map Input", value=st.session_state.col_map, label_visibility="collapsed", help="ColIndex:Code", height=100)
        c1, c2 = st.columns(2)
        if c1.button("Clear Map"): st.session_state.col_map = ""; st.rerun()
        if c2.button("Restore Map"): st.session_state.col_map = DEFAULTS["col_map"]; st.rerun()

    with col_set3:
        st.subheader("Agency & Column-wise FG Override")
        st.session_state.agency_override = st.text_area("Agency Override Input", value=st.session_state.agency_override, label_visibility="collapsed", help="Agency:ColIndex:CustomFG", height=100)
        c1, c2 = st.columns(2)
        if c1.button("Clear Override"): st.session_state.agency_override = ""; st.rerun()
        if c2.button("Restore Override"): st.session_state.agency_override = DEFAULTS["agency_override"]; st.rerun()

    st.markdown("---")
    col_set4, col_set5 = st.columns(2)
    with col_set4:
        st.subheader("📧 Email Dispatch Settings")
        st.session_state.email_user = st.text_input("Sender Email ID", value=st.session_state.email_user)
        st.session_state.email_pass = st.text_input("Email App Password", type="password", value=st.session_state.email_pass)
        st.session_state.recipient = st.text_input("Recipient Email", value=st.session_state.recipient)
    with col_set5:
        st.subheader("📱 WhatsApp Notification")
        st.session_state.whatsapp = st.text_input("WhatsApp Number (e.g., 919876543210)", value=st.session_state.whatsapp)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Reset All Settings to Defaults"):
            for k, v in DEFAULTS.items():
                if k != "selected_theme":
                    st.session_state[k] = v
            st.rerun()

default_fg_code = st.session_state.fg_code
col_mapping_input = st.session_state.col_map
agency_fg_override = st.session_state.agency_override
default_fallback_route = st.session_state.route
email_user = st.session_state.email_user
email_pass = st.session_state.email_pass
recipient_email = st.session_state.recipient
whatsapp_num = st.session_state.whatsapp
vehicle_capacity_limit = st.session_state.vehicle_capacity
default_vehicle_no = st.session_state.vehicle_no
default_driver_mobile = st.session_state.driver_mobile

direct_col_mapping = {}
for line in col_mapping_input.split('\n'):
    if ':' in line:
        parts = line.split(':')
        idx_str = parts[0].strip()
        if idx_str.isdigit():
            direct_col_mapping[int(idx_str)] = parts[1].strip()

agency_col_override_map = {}
for line in agency_fg_override.split('\n'):
    parts = line.split(':')
    if len(parts) == 3:
        ag, col_idx, fg = parts[0].strip(), parts[1].strip(), parts[2].strip()
        if ag.isdigit() and col_idx.isdigit():
            agency_col_override_map[(int(ag), int(col_idx))] = fg

st.title(f"💼 Enterprise Sales Order & Dispatch Hub ({st.session_state.selected_theme})")
st.markdown("Upload inbound demand files to execute **Vehicle Capacity Clubbing**, **Demand vs Dispatch vs Pending Tracking**, and generate print-ready delivery schedules.")
st.markdown("---")

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

if uploaded_inputs:
    with st.expander("🔍 Pre-flight File Health Check Report", expanded=False):
        preflight_logs = []
        for uploaded_file in uploaded_inputs:
            short_filename = uploaded_file.name
            if short_filename.lower() == "output.xlsx":
                continue
            try:
                df_prev = pd.read_excel(io.BytesIO(uploaded_file.getvalue()), header=None)
                fg_found = any("FG" in str(df_prev.iloc[r, c]).strip().upper() for r in range(df_prev.shape[0]) for c in range(df_prev.shape[1]))
                if fg_found:
                    preflight_logs.append({"File Name": short_filename, "Health Status": "🟢 Healthy", "Details": "FG Header detected successfully"})
                else:
                    preflight_logs.append({"File Name": short_filename, "Health Status": "🔴 Warning", "Details": "'FG' header missing"})
            except Exception as e:
                preflight_logs.append({"File Name": short_filename, "Health Status": "❌ Corrupt", "Details": str(e)})
        if preflight_logs:
            st.dataframe(pd.DataFrame(preflight_logs), use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 5. CORE BATCH PROCESSING & DISPATCH ENGINE
# ==========================================
if st.button("🚀 Process Batch Orders & Generate Dispatch Plan", type="primary"):
    if uploaded_inputs:
        st.session_state.processed_files = []
        st.session_state.comparison_summary = []
        st.session_state.skipped_rows_log = []
        st.session_state.anomaly_logs = []
        st.session_state.unmapped_current_batch = []
        
        total_input_qty = 0
        total_gen_qty = 0
        total_valid_orders = 0
        total_missing_orders = 0
        total_skipped_rows = 0
        
        db_records_to_insert = []
        unmapped_records_to_insert = []
        output_files_to_store = []
        traceability_records = []
        discrepancy_records = []
        dispatch_plan_records = []
        
        with st.spinner("⚡ Running Vehicle Capacity Clubbing & Shortfall Pending Calculations... Please wait."):
            try:
                try:
                    with open("Output.xlsx", "rb") as f:
                        template_bytes = f.read()
                except FileNotFoundError:
                    st.error("❌ 'Output.xlsx' template file repository mein nahi mili.")
                    st.stop()
                
                conn_stk = sqlite3.connect("sales_history.db")
                df_stock_db = pd.read_sql("SELECT * FROM inventory_stock_table", conn_stk)
                conn_stk.close()
                stock_inventory_dict = dict(zip(df_stock_db['fg_code'], df_stock_db['available_qty'])) if not df_stock_db.empty else {}

                ist_now = get_ist_now()
                today_date = ist_now.strftime("%Y-%m-%d")
                timestamp = ist_now.strftime("%H%M%S")
                batch_ts = ist_now.strftime("%Y-%m-%d %H:%M:%S")

                for uploaded_file in uploaded_inputs:
                    short_filename = uploaded_file.name
                    if short_filename.lower() == "output.xlsx":
                        continue

                    file_bytes = uploaded_file.getvalue()
                    df_input = pd.read_excel(io.BytesIO(file_bytes), header=None)

                    fg_row, fg_col = -1, -1
                    for r in range(df_input.shape[0]):
                        for c in range(df_input.shape[1]):
                            val = str(df_input.iloc[r, c]).strip().upper()
                            if "FG" in val:
                                fg_row, fg_col = r, c
                                break
                        if fg_row != -1:
                            break

                    if fg_row == -1:
                        st.warning(f"⚠️ '{short_filename}' mein 'FG' header nahi mila. File skip ho rahi hai.")
                        continue

                    total_col = df_input.shape[1]
                    for cSearch in range(fg_col, df_input.shape[1]):
                        is_total = False
                        for scan_r in range(max(0, fg_row - 10), min(fg_row + 3, df_input.shape[0])):
                            cell_val = str(df_input.iloc[scan_r, cSearch]).strip().upper()
                            if any(kw in cell_val for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                                is_total = True
                                break
                            if scan_r >= fg_row + 1 and ("SUM" in cell_val or "=" in cell_val):
                                is_total = True
                                break
                        if is_total:
                            total_col = cSearch
                            break

                    route_num = default_fallback_route if default_fallback_route != "" else "22"
                    ignore_list = ["RT", "DR", "RT DR", "ROUTE", "SALES PERSON", "CONTACT NO:", "MATERIAL CODE"]
                    
                    for r in range(fg_row):
                        for c in range(min(total_col, 30)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if upper_val in ignore_list:
                                continue
                            is_product_code = any(upper_val.startswith(p) for p in ["PC", "MS", "M", "GM", "DP", "SKU", "FG"])
                            if is_product_code:
                                continue
                            if cell_val != "" and 1 <= len(cell_val) <= 3:
                                if any(char.isdigit() for char in cell_val):
                                    route_num = cell_val
                                    break
                        if route_num != (default_fallback_route if default_fallback_route != "" else "22"):
                            break

                    if default_fallback_route != "" and default_fallback_route != "22":
                        route_num = default_fallback_route

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    agency_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        valid_count = 0
                        for rCheck in range(fg_row + 1, min(fg_row + 15, df_input.shape[0])):
                            v = df_input.iloc[rCheck, cSearch]
                            if pd.notna(v):
                                s_val = str(v).replace('.0', '').strip()
                                if s_val.isdigit() and 1 <= len(s_val) <= 5:
                                    valid_count += 1
                        if valid_count >= 3:
                            agency_col = cSearch
                            break

                    if agency_col == -1 and fg_col > 0:
                        agency_col = fg_col - 1

                    dr_code_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        sample_val = str(df_input.iloc[fg_row + 1, cSearch] if fg_row + 1 < df_input.shape[0] else "").strip().upper()
                        if re.match(r'^DR\d+', sample_val):
                            dr_code_col = cSearch
                            break
                        matched_count = 0
                        for offset in range(1, min(4, df_input.shape[0] - fg_row)):
                            v = str(df_input.iloc[fg_row + offset, cSearch]).strip().upper()
                            if re.match(r'^DR\d+', v):
                                matched_count += 1
                        if matched_count > 0:
                            dr_code_col = cSearch
                            break

                    valid_cols = []
                    for c in range(fg_col, total_col):
                        fg_code_header = str(df_input.iloc[fg_row, c] if fg_row >= 0 else "").strip()
                        if any(kw in fg_code_header.upper() for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                            break
                        valid_cols.append((c, fg_code_header))

                    wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_valid = wb_valid["Order Data"] if "Order Data" in wb_valid.sheetnames else wb_valid.active

                    wb_missing = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_missing = wb_missing["Order Data"] if "Order Data" in wb_missing.sheetnames else wb_missing.active

                    valid_row, missing_row = 6, 6
                    valid_order_num, missing_order_num = 1, 1
                    agency_counts_valid, agency_counts_missing = {}, {}
                    valid_items_created, missing_items_created = 0, 0
                    file_comparison_rows = []
                    file_input_qty = 0

                    current_vehicle_load = 0.0
                    assigned_vehicle_no = default_vehicle_no

                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        if pd.isna(agency) or str(agency).strip() in ["", "nan", "None"]:
                            continue
                        
                        agency_str = str(agency).replace('.0','').strip()
                        if any(kw in agency_str.upper() for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                            continue

                        if not agency_str.isdigit() or not (1 <= len(agency_str) <= 5):
                            st.session_state.skipped_rows_log.append({
                                "File Name": short_filename,
                                "Row Index": r + 1,
                                "Agency Value": str(agency),
                                "Reason": "Invalid or Non-numeric Agency Number / Summary Row"
                            })
                            total_skipped_rows += 1
                            continue

                        agency_val = int(agency_str)
                        
                        valid_row_quantities = []
                        row_total_qty = 0
                        for c, fg_code_header in valid_cols:
                            if c >= total_col:
                                continue
                            sku_qty = df_input.iloc[r, c]
                            if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                try:
                                    qty_val = float(sku_qty)
                                    if qty_val > 0:
                                        row_total_qty += qty_val
                                        file_input_qty += qty_val
                                        
                                        # CORRECT FG CODE RESOLUTION
                                        current_fg_code = agency_col_override_map.get((agency_val, c), direct_col_mapping.get(c, fg_code_header if fg_code_header and str(fg_code_header).upper() != "NAN" else default_fg_code))
                                        if not current_fg_code or str(current_fg_code).strip() == "" or str(current_fg_code).upper() == "NAN":
                                            current_fg_code = default_fg_code

                                        avail_stock = stock_inventory_dict.get(current_fg_code, 99999.0)
                                        
                                        if avail_stock >= qty_val:
                                            dispatched_qty = qty_val
                                            pending_qty = 0.0
                                            stock_inventory_dict[current_fg_code] = avail_stock - qty_val
                                        else:
                                            dispatched_qty = avail_stock if avail_stock > 0 else 0.0
                                            pending_qty = qty_val - dispatched_qty
                                            stock_inventory_dict[current_fg_code] = 0.0

                                        valid_row_quantities.append((c, current_fg_code, qty_val, dispatched_qty, pending_qty))
                                except ValueError:
                                    pass

                        if not valid_row_quantities:
                            continue

                        row_dispatch_sum = sum([item[3] for item in valid_row_quantities])
                        row_pending_sum = sum([item[4] for item in valid_row_quantities])
                        
                        status_str = "Fully Dispatched"
                        if current_vehicle_load + row_dispatch_sum > vehicle_capacity_limit:
                            allowed_load = max(0.0, vehicle_capacity_limit - current_vehicle_load)
                            if allowed_load > 0:
                                status_str = "Partially Dispatched (Vehicle Capacity Split)"
                            else:
                                status_str = "Pending Next Day (Vehicle Full)"
                                row_pending_sum += row_dispatch_sum
                                row_dispatch_sum = 0.0

                        current_vehicle_load += row_dispatch_sum
                        if current_vehicle_load >= vehicle_capacity_limit:
                            current_vehicle_load = row_dispatch_sum

                        for _, fgc, d_dem, d_disp, d_pend in valid_row_quantities:
                            dispatch_plan_records.append((today_date, str(route_num), assigned_vehicle_no, default_driver_mobile, str(agency_val), fgc, d_dem, d_disp, d_pend, status_str, batch_ts))

                        has_dr_code = False
                        clean_dr = ""
                        if dr_code_col >= 0 and dr_code_col < df_input.shape[1]:
                            raw_dr = df_input.iloc[r, dr_code_col]
                            if pd.notna(raw_dr):
                                val_str = str(raw_dr).replace('.0', '').strip()
                                if "DR" in val_str.upper() and val_str.upper() != "0":
                                    has_dr_code = True
                                    clean_dr = val_str

                        if not has_dr_code:
                            for c_scan in range(fg_col):
                                cell_val = df_input.iloc[r, c_scan]
                                if pd.notna(cell_val):
                                    val_str = str(cell_val).replace('.0', '').strip()
                                    if "DR" in val_str.upper() and val_str.upper() != "0":
                                        has_dr_code = True
                                        clean_dr = val_str
                                        break

                        if not has_dr_code:
                            conn_lookup = sqlite3.connect("sales_history.db")
                            cursor_lookup = conn_lookup.cursor()
                            cursor_lookup.execute("""
                                SELECT dr_code FROM unique_routes_master 
                                WHERE route_no = ? AND agency_no = ? AND dr_code LIKE 'DR%' 
                                LIMIT 1
                            """, (str(route_num), str(agency_val)))
                            db_match = cursor_lookup.fetchone()
                            conn_lookup.close()
                            
                            if db_match:
                                has_dr_code = True
                                clean_dr = db_match[0]

                        if not has_dr_code:
                            unmapped_record = (short_filename, str(route_num), str(agency_val), f"NEW_CUST_{agency_val}", ist_now.strftime("%Y-%m-%d %H:%M:%S"))
                            if unmapped_record not in unmapped_records_to_insert:
                                unmapped_records_to_insert.append(unmapped_record)
                            
                            current_unmapped_dict = {
                                "File Name": short_filename,
                                "Route": str(route_num),
                                "Agency": agency_val,
                                "Status": "Generated via NEW_CUST (Missing DR in File and Master DB)"
                            }
                            if current_unmapped_dict not in st.session_state.unmapped_current_batch:
                                st.session_state.unmapped_current_batch.append(current_unmapped_dict)

                        final_dr = clean_dr if has_dr_code else f"NEW_CUST_{agency_val}"
                        
                        if has_dr_code and clean_dr.startswith("DR"):
                            db_record = (short_filename, str(route_num), str(agency_val), str(clean_dr), ist_now.strftime("%Y-%m-%d %H:%M:%S"))
                            if db_record not in db_records_to_insert:
                                db_records_to_insert.append(db_record)

                        if row_total_qty > 500:
                            st.session_state.anomaly_logs.append({
                                "File Name": short_filename,
                                "Agency": agency_val,
                                "Route": route_num,
                                "Total Qty": row_total_qty,
                                "Flag": "⚠️ High Volume Spike (>500)"
                            })

                        if has_dr_code:
                            agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                            current_seq = agency_counts_valid[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_valid, valid_row, valid_order_num, clean_dr, "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            current_seq = agency_counts_missing[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-NEW-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_missing, missing_row, missing_order_num, f"NEW_CUST_{agency_val}", "Missing DR"

                        item_id = 10
                        for c, current_fg, q_in, q_disp, q_pend in valid_row_quantities:
                            if q_disp <= 0:
                                continue

                            total_input_qty += q_in
                            total_gen_qty += q_disp
                            
                            file_comparison_rows.append({
                                "File Name": short_filename,
                                "Status": file_category,
                                "Agency": agency_val,
                                "DR Code": dr_to_use,
                                "FG Code": current_fg,
                                "Input Qty": q_in,
                                "Generated Qty": q_disp
                            })
                            
                            diff_val = q_in - q_disp
                            if diff_val != 0:
                                discrepancy_records.append((batch_ts, short_filename, str(agency_val), dr_to_use, current_fg, q_in, q_disp, diff_val, batch_ts))

                            target_ws.cell(row=current_r, column=2, value=order_num)
                            target_ws.cell(row=current_r, column=3, value="OR")
                            target_ws.cell(row=current_r, column=4, value="SO20")
                            target_ws.cell(row=current_r, column=5, value=10)
                            target_ws.cell(row=current_r, column=6, value=20)
                            target_ws.cell(row=current_r, column=7, value=dr_to_use)
                            target_ws.cell(row=current_r, column=8, value=dr_to_use)
                            target_ws.cell(row=current_r, column=9, value=ref_number)
                            target_ws.cell(row=current_r, column=10, value=today_date)
                            target_ws.cell(row=current_r, column=11, value=today_date)
                            target_ws.cell(row=current_r, column=15, value=item_id)
                            target_ws.cell(row=current_r, column=16, value=current_fg)
                            target_ws.cell(row=current_r, column=19, value=q_disp)
                            target_ws.cell(row=current_r, column=20, value="Bag")
                            target_ws.cell(row=current_r, column=22, value=2100)
                            target_ws.cell(row=current_r, column=26, value=str(route_num))
                            target_ws.cell(row=current_r, column=27, value=agency_val)
                            
                            item_id += 10
                            current_r += 1

                        if has_dr_code:
                            valid_row, valid_order_num, valid_items_created, total_valid_orders = current_r, valid_order_num + 1, valid_items_created + 1, total_valid_orders + 1
                        else:
                            missing_row, missing_order_num, missing_items_created, total_missing_orders = current_r, missing_order_num + 1, missing_items_created + 1, total_missing_orders + 1

                    if valid_items_created > 0:
                        buf_valid = io.BytesIO()
                        wb_valid.save(buf_valid)
                        buf_valid.seek(0)
                        out_fname = safe_route_num + "_" + today_date + "_" + timestamp + "_Valid.xlsx"
                        st.session_state.processed_files.append({
                            "name": short_filename + " (Valid DR)",
                            "data": buf_valid.getvalue(),
                            "filename": out_fname,
                            "orders": valid_items_created
                        })
                        output_files_to_store.append((out_fname, "Valid DR", buf_valid.getvalue(), ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname, "Valid DR", 1, batch_ts))

                    if missing_items_created > 0:
                        buf_missing = io.BytesIO()
                        wb_missing.save(buf_missing)
                        buf_missing.seek(0)
                        out_fname_miss = safe_route_num + "_" + today_date + "_" + timestamp + "_Missing_DR.xlsx"
                        st.session_state.processed_files.append({
                            "name": short_filename + " (Missing DR / New)",
                            "data": buf_missing.getvalue(),
                            "filename": out_fname_miss,
                            "orders": missing_items_created
                        })
                        output_files_to_store.append((out_fname_miss, "Missing DR", buf_missing.getvalue(), ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname_miss, "Missing DR", 1, batch_ts))

                    if file_comparison_rows:
                        df_comp = pd.DataFrame(file_comparison_rows)
                        df_pivot = df_comp.pivot_table(
                            index=["File Name", "Status", "Agency", "DR Code", "FG Code"],
                            values=["Input Qty", "Generated Qty"],
                            aggfunc="sum"
                        ).reset_index()
                        df_pivot["Difference"] = df_pivot["Input Qty"] - df_pivot["Generated Qty"]
                        st.session_state.comparison_summary.append(df_pivot)

                # --- Update Master DB, Unmapped Ledger & Output Files Ledger ---
                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.executemany("""
                    INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at)
                    VALUES (?, ?, ?, ?, ?)
                """, db_records_to_insert)

                cursor.executemany("""
                    INSERT OR IGNORE INTO unmapped_missing_dr_ledger (file_name, route_no, agency_no, dr_code, created_at)
                    VALUES (?, ?, ?, ?, ?)
                """, unmapped_records_to_insert)
                
                for fname, ftype, fdata, fdate in output_files_to_store:
                    cursor.execute("""
                        INSERT OR IGNORE INTO output_files_ledger (file_name, file_type, file_data, created_at)
                        VALUES (?, ?, ?, ?)
                    """, (fname, ftype, fdata, fdate))

                cursor.executemany("""
                    INSERT INTO input_output_traceability (batch_timestamp, input_file_name, input_file_blob, total_input_qty, generated_output_file, output_type, version_no, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, traceability_records)

                if discrepancy_records:
                    cursor.executemany("""
                        INSERT INTO discrepancy_audit_ledger (batch_timestamp, file_name, agency_no, dr_code, fg_code, input_qty, generated_qty, difference, logged_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, discrepancy_records)

                if dispatch_plan_records:
                    cursor.executemany("""
                        INSERT INTO dispatch_planning_ledger (dispatch_date, route_no, vehicle_no, driver_mobile, agency_no, fg_code, demand_qty, dispatched_qty, pending_qty, status, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, dispatch_plan_records)
                
                cursor.execute(
                    "INSERT INTO history_logs (timestamp, files_count, total_qty, status) VALUES (?, ?, ?, ?)",
                    (get_ist_now().strftime("%Y-%m-%d %H:%M:%S"), len(uploaded_inputs), total_input_qty, "Success")
                )
                conn.commit()
                conn.close()

                st.session_state.kpi_data = {
                    "input_qty": total_input_qty,
                    "gen_qty": total_gen_qty,
                    "valid_count": total_valid_orders,
                    "missing_count": total_missing_orders,
                    "skipped_count": total_skipped_rows
                }

                st.success("✅ Batch Processing, Vehicle Capacity Planning & Audit Ledgers Updated Successfully!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya pehle demand files upload karein!")

# ==========================================
# 6. DASHBOARD, KPI, ANALYTICS & EXPORTS
# ==========================================
if st.session_state.processed_files or st.session_state.skipped_rows_log:
    st.markdown("---")
    st.markdown("### 📈 Batch Performance & KPI Summary")
    kpi = st.session_state.kpi_data
    
    total_processed_orders = kpi['valid_count'] + kpi['missing_count']
    success_rate = (kpi['valid_count'] / total_processed_orders * 100) if total_processed_orders > 0 else 0
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Input Qty", f"{kpi['input_qty']:,.0f}")
    col2.metric("Generated Qty", f"{kpi['gen_qty']:,.0f}")
    col3.metric("Valid Orders", kpi['valid_count'])
    col4.metric("Success Rate", f"{success_rate:.1f}%")
    col5.metric("Skipped Rows", kpi['skipped_count'], delta_color="inverse")

    if kpi['skipped_count'] > 5:
        st.warning(f"⚠️ **Smart Audit Alert:** {kpi['skipped_count']} rows skipped check exception logs.")

    # --- AI Sales Demand Forecasting & Velocity Health Scorecard ---
    st.markdown("---")
    st.markdown("### 🤖 AI Sales Demand Forecasting & Velocity Health Scorecard")
    
    demand_health_score = success_rate
    forecast_confidence = "🟢 High Confidence (Stable Batch Flow)" if demand_health_score >= 90 else ("🟡 Moderate Risk (Unmapped Fallbacks Detected)" if demand_health_score >= 70 else "🔴 Critical Review Needed (High Missing DR Ratio)")
    
    f_col1, f_col2, f_col3 = st.columns(3)
    f_col1.metric("Batch Demand Health Score", f"{demand_health_score:.1f}%", delta="Optimal Flow" if demand_health_score >= 90 else "Attention Needed")
    f_col2.metric("Forecast Confidence Status", forecast_confidence)
    f_col3.metric("Projected Next-Cycle Qty", f"{kpi['gen_qty'] * 1.05:,.0f} Units", delta="+5% Trend")

    # --- AI Anomaly & Unmapped Missing DR Alerts ---
    if st.session_state.anomaly_logs:
        st.markdown("---")
        st.markdown("### 🤖 AI Demand Spike & Anomaly Detector Alerts")
        st.warning("⚠️ System detected high-volume demand spikes (>500 units) in the following agencies:")
        st.dataframe(pd.DataFrame(st.session_state.anomaly_logs), use_container_width=True)

    if st.session_state.unmapped_current_batch:
        st.markdown("---")
        st.markdown("### 🚨 Unmapped Missing DRs Alert List")
        st.error("⚠️ The following agencies had valid quantity (>0) but no DR code in file or Master DB. They were successfully processed using `NEW_CUST` fallback and logged into the Unmapped Ledger:")
        st.dataframe(pd.DataFrame(st.session_state.unmapped_current_batch), use_container_width=True)

    # --- ADVANCED TABBED VISUAL ANALYTICS ---
    if st.session_state.comparison_summary:
        st.markdown("---")
        st.markdown("### 📊 Advanced Visual Analytics Dashboard")
        combined_df_chart = pd.concat(st.session_state.comparison_summary, ignore_index=True)
        
        tab1, tab2 = st.tabs(["📊 Agency-wise Breakdown", "📦 SKU-wise Share"])
        with tab1:
            st.bar_chart(combined_df_chart.groupby("Agency")["Generated Qty"].sum())
        with tab2:
            st.bar_chart(combined_df_chart.groupby("FG Code")["Generated Qty"].sum())

    st.markdown("---")
    st.markdown("### 📥 Bulk Download & Notifications")
    
    # --- Advanced Email Config Expander ---
    with st.expander("✉️ Advanced Email Dispatch Options (Custom Subject & Note)"):
        email_subject_custom = st.text_input("Custom Email Subject Line", f"🚀 Sales Orders Batch Execution Report (IST) - {get_ist_now().strftime('%Y-%m-%d')}")
        email_notes_custom = st.text_area("Custom Remarks / Notes to Include in Email Body", "All routes verified and processed successfully.")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])
    
    col_zip, col_pdf, col_summary, col_json, col_print, col_email, col_wa = st.columns(7)
    
    with col_zip:
        st.download_button(
            label="📦 ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"Batch_Orders_{get_ist_now().strftime('%Y-%m-%d')}.zip",
            mime="application/zip",
            key="zip_download"
        )
        
    with col_pdf:
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(190, 10, "Enterprise Sales Order Summary Invoice", ln=True, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(190, 6, f"Generated On (IST): {get_ist_now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
            pdf.ln(10)
            
            pdf.set_font("Arial", "B", 11)
            pdf.cell(100, 8, "Metric Description", border=1)
            pdf.cell(90, 8, "Value", border=1, ln=True)
            
            pdf.set_font("Arial", "", 11)
            metrics_list = [
                ("Total Input Quantity", f"{kpi['input_qty']:,.0f}"),
                ("Total Generated Quantity", f"{kpi['gen_qty']:,.0f}"),
                ("Valid DR Orders", str(kpi['valid_count'])),
                ("Missing DR Orders (NEW_CUST)", str(kpi['missing_count'])),
                ("Skipped Rows Logged", str(kpi['skipped_count'])),
                ("Success Rate", f"{success_rate:.1f}%"),
                ("Demand Health Score", f"{demand_health_score:.1f}%")
            ]
            for m_desc, m_val in metrics_list:
                pdf.cell(100, 8, m_desc, border=1)
                pdf.cell(90, 8, m_val, border=1, ln=True)
                
            pdf_bytes = bytes(pdf.output())
            st.download_button(
                label="📄 PDF",
                data=pdf_bytes,
                file_name=f"Sales_Invoice_{get_ist_now().strftime('%Y-%m-%d')}.pdf",
                mime="application/pdf",
                key="pdf_invoice_download"
            )
        except Exception as e:
            st.error(f"PDF Error: {str(e)}")

    with col_summary:
        summary_txt = f"""=== ENTERPRISE SALES ORDER SUMMARY (IST) ===
Date/Time: {get_ist_now().strftime('%Y-%m-%d %H:%M:%S')}
----------------------------------------
Total Input Quantity : {kpi['input_qty']:,.0f}
Total Generated Qty  : {kpi['gen_qty']:,.0f}
Valid DR Orders      : {kpi['valid_count']}
Missing DR Orders    : {kpi['missing_count']}
Skipped Rows Logged  : {kpi['skipped_count']}
Success Rate         : {success_rate:.1f}%
Demand Health Score  : {demand_health_score:.1f}%
----------------------------------------
Generated Files Count: {len(st.session_state.processed_files)}
Status: Successfully Processed & Audited
========================================"""
        st.download_button(
            label="📄 TXT",
            data=summary_txt.encode('utf-8'),
            file_name=f"Summary_Report_{get_ist_now().strftime('%Y-%m-%d')}.txt",
            mime="text/plain",
            key="summary_txt_download"
        )
        
    with col_json:
        json_data = json.dumps({
            "timestamp": get_ist_now().strftime('%Y-%m-%d %H:%M:%S'),
            "metrics": kpi,
            "success_rate": f"{success_rate:.1f}%",
            "demand_health_score": f"{demand_health_score:.1f}%"
        }, indent=4)
        st.download_button(
            label="💾 JSON",
            data=json_data.encode('utf-8'),
            file_name=f"Audit_Backup_{get_ist_now().strftime('%Y-%m-%d')}.json",
            mime="application/json",
            key="json_backup_download"
        )
        
    with col_print:
        print_html = """
        <div style="width:100%; margin:0; padding:0;">
            <button onclick="parent.window.print()" style="width:100%; height:38px; background:#2563eb; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer; font-family:sans-serif; display:flex; align-items:center; justify-content:center;">
                🖨️ Print
            </button>
        </div>
        """
        components.html(print_html, height=50)
        
    with col_email:
        if st.button("📧 Email"):
            if email_user and email_pass and recipient_email:
                try:
                    conn = sqlite3.connect("sales_history.db")
                    df_master_email = pd.read_sql("SELECT * FROM unique_routes_master", conn)
                    conn.close()
                    
                    excel_buffer = io.BytesIO()
                    df_master_email.to_excel(excel_buffer, index=False, sheet_name="Master Routes")
                    excel_buffer.seek(0)

                    unmapped_email_html = ""
                    if st.session_state.unmapped_current_batch:
                        unmapped_email_html = "<h3 style='color: #dc2626;'>🚨 Unmapped Missing DRs Alert List</h3><table style='border-collapse: collapse; width: 100%; margin-top: 10px; border-radius: 6px; overflow: hidden;'><tr style='background-color: #ef4444; color: white;'><th style='padding: 8px; text-align: left;'>File Name</th><th style='padding: 8px; text-align: left;'>Route</th><th style='padding: 8px; text-align: left;'>Agency</th><th style='padding: 8px; text-align: left;'>Status</th></tr>"
                        for item in st.session_state.unmapped_current_batch:
                            unmapped_email_html += f"<tr style='background-color: #fef2f2;'><td style='padding: 8px; border-bottom: 1px solid #fee2e2;'>{item['File Name']}</td><td style='padding: 8px; border-bottom: 1px solid #fee2e2;'>{item['Route']}</td><td style='padding: 8px; border-bottom: 1px solid #fee2e2;'>{item['Agency']}</td><td style='padding: 8px; border-bottom: 1px solid #fee2e2; color: #dc2626;'>{item['Status']}</td></tr>"
                        unmapped_email_html += "</table>"

                    msg = EmailMessage()
                    msg['Subject'] = email_subject_custom
                    msg['From'] = email_user
                    msg['To'] = recipient_email
                    
                    html_content = f"""
                    <html>
                      <body style="font-family: Arial, sans-serif; color: #333; background-color: #f9fafb; padding: 20px;">
                        <div style="max-width: 600px; background: #ffffff; padding: 25px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                          <h2 style="color: #10b981; border-bottom: 2px solid #e5e7eb; padding-bottom: 10px;">📊 Sales Order Batch Automation Hub</h2>
                          <p>Hello Team,</p>
                          <p>The daily inbound demand batch has been processed successfully on <b>{get_ist_now().strftime('%Y-%m-%d %H:%M:%S')} IST</b>.</p>
                          <p style="background: #f0fdf4; border-left: 4px solid #10b981; padding: 10px; margin: 15px 0;"><b>Remarks:</b> {email_notes_custom}</p>
                          <table style="border-collapse: collapse; width: 100%; margin-top: 15px; border-radius: 6px; overflow: hidden;">
                            <tr style="background-color: #10b981; color: white;">
                              <th style="padding: 10px; text-align: left;">Metric</th>
                              <th style="padding: 10px; text-align: left;">Value</th>
                            </tr>
                            <tr style="background-color: #f3f4f6;">
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Total Input Qty</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;"><b>{kpi['input_qty']:,.0f}</b></td>
                            </tr>
                            <tr>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Valid Orders</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">{kpi['valid_count']}</td>
                            </tr>
                            <tr style="background-color: #f3f4f6;">
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Missing DR Orders (NEW_CUST)</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;"><b>{kpi['missing_count']}</b></td>
                            </tr>
                            <tr>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Demand Health Score</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;"><b>{demand_health_score:.1f}%</b></td>
                            </tr>
                          </table>
                          {unmapped_email_html}
                          <p style="margin-top: 25px; color: #666; font-size: 12px; border-top: 1px solid #e5e7eb; paddingTop: 10px;">Master Route-Agency-DR Database attached herewith.</p>
                        </div>
                      </body>
                    </html>
                    """
                    msg.set_content("Please enable HTML to view this report.")
                    msg.add_alternative(html_content, subtype='html')
                    
                    for item in st.session_state.processed_files:
                        msg.add_attachment(item['data'], maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=item['filename'])
                    
                    msg.add_attachment(excel_buffer.getvalue(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=f"Unique_Routes_Master_{get_ist_now().strftime('%Y-%m-%d')}.xlsx")
                    
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email_user, email_pass)
                        smtp.send_message(msg)
                    st.success("✅ Email dispatched with Master DB attached!")
                except Exception as e:
                    st.error(f"❌ Email failed: {str(e)}")
            else:
                st.warning("⚠️ Enter email credentials!")

    with col_wa:
        if whatsapp_num:
            wa_text = f"Sales Order Batch Ready! Total Qty: {kpi['input_qty']}, Health Score: {demand_health_score:.1f}%."
            wa_link = f"https://wa.me/{whatsapp_num}?text={urllib.parse.quote(wa_text)}"
            st.markdown(f'<a href="{wa_link}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer; display:flex; align-items:center; justify-content:center;">📱 WhatsApp</button></a>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("##### Individual File Downloads:")
    for i, item in enumerate(st.session_state.processed_files):
        if st.download_button(
            label=f"📥 Download {item['name']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_file_{i}_{item['filename']}"
        ):
            st.toast(f"🎉 '{item['filename']}' downloaded!", icon="📥")

# ==========================================
# 7. MULTI-DATABASE MANAGEMENT & UNIVERSAL COLUMN/ROW MODIFIER
# ==========================================
st.markdown("---")
with st.expander("🗄️ View, Export & Manage All Databases (Add/Delete Columns & Rows)"):
    st.markdown("Yahan aap saare databases ke records dekh sakte hain, naye columns add kar sakte hain, rows delete kar sakte hain, aur tables wipe kar sakte hain.")
    try:
        conn = sqlite3.connect("sales_history.db")
        df_master = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
        df_unmapped = pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger ORDER BY id DESC", conn)
        df_outputs = pd.read_sql("SELECT id, file_name, file_type, created_at FROM output_files_ledger ORDER BY id DESC", conn)
        df_trace = pd.read_sql("SELECT id, batch_timestamp, input_file_name, total_input_qty, generated_output_file, output_type, version_no FROM input_output_traceability ORDER BY id DESC", conn)
        df_audit = pd.read_sql("SELECT * FROM discrepancy_audit_ledger ORDER BY id DESC", conn)
        df_dispatch = pd.read_sql("SELECT * FROM dispatch_planning_ledger ORDER BY id DESC", conn)
        df_stock = pd.read_sql("SELECT * FROM inventory_stock_table ORDER BY id DESC", conn)
        conn.close()
        
        tab_m1, tab_m2, tab_m3, tab_m4, tab_m5, tab_m6, tab_m7, tab_m8, tab_m9 = st.tabs([
            "📋 Master", "🚨 Unmapped", "📦 Outputs", "🚚 Dispatch", "🔗 Trace", "🔍 Audit", "📅 Monthly", "⏳ Pending", "🛠️ DB Schema Editor"
        ])
        
        def render_table_manager(table_name, df_data):
            st.dataframe(df_data, use_container_width=True)
            col_ed1, col_ed2, col_ed3 = st.columns(3)
            with col_ed1:
                new_col_name = st.text_input(f"New Column Name ({table_name})", key=f"col_name_{table_name}")
                if st.button(f"➕ Add Column to {table_name}", key=f"btn_add_col_{table_name}"):
                    if new_col_name:
                        try:
                            conn_c = sqlite3.connect("sales_history.db")
                            cur_c = conn_c.cursor()
                            cur_c.execute(f"ALTER TABLE {table_name} ADD COLUMN {new_col_name} TEXT")
                            conn_c.commit()
                            conn_c.close()
                            st.success(f"✅ Column '{new_col_name}' added successfully to {table_name}!")
                            st.rerun()
                        except Exception as ex:
                            st.error(f"Error: {str(ex)}")
            with col_ed2:
                row_id_rem = st.number_input(f"Record ID to Delete ({table_name})", min_value=1, step=1, key=f"del_id_{table_name}")
                if st.button(f"🗑️ Delete Row & Reset ID", key=f"btn_del_row_{table_name}"):
                    try:
                        conn_r = sqlite3.connect("sales_history.db")
                        cur_r = conn_r.cursor()
                        cur_r.execute(f"DELETE FROM {table_name} WHERE id = ?", (row_id_rem,))
                        cur_r.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}'")
                        conn_r.commit()
                        conn_r.close()
                        st.success(f"✅ Record ID {row_id_rem} deleted from {table_name}!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error: {str(ex)}")
            with col_ed3:
                if st.button(f"🚨 Wipe Entire {table_name}", key=f"btn_wipe_{table_name}", type="secondary"):
                    try:
                        conn_w = sqlite3.connect("sales_history.db")
                        cur_w = conn_w.cursor()
                        cur_w.execute(f"DELETE FROM {table_name}")
                        cur_w.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}'")
                        conn_w.commit()
                        conn_w.close()
                        st.success(f"✅ Table {table_name} wiped successfully!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Error: {str(ex)}")

        with tab_m1:
            render_table_manager("unique_routes_master", df_master)
        with tab_m2:
            render_table_manager("unmapped_missing_dr_ledger", df_unmapped)
        with tab_m3:
            render_table_manager("output_files_ledger", df_outputs)
        with tab_m4:
            render_table_manager("dispatch_planning_ledger", df_dispatch)
        with tab_m5:
            render_table_manager("input_output_traceability", df_trace)
        with tab_m6:
            render_table_manager("discrepancy_audit_ledger", df_audit)
        with tab_m7:
            if not df_dispatch.empty:
                st.dataframe(df_dispatch, use_container_width=True)
                monthly_buf = io.BytesIO()
                df_dispatch.to_excel(monthly_buf, index=False, sheet_name="Monthly Dispatch Master")
                monthly_buf.seek(0)
                st.download_button("📥 Export Monthly Dispatch Masterfile (.xlsx)", monthly_buf.getvalue(), f"Monthly_Dispatch_{get_ist_now().strftime('%Y-%m')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_monthly")
            else:
                st.info("No records.")
        with tab_m8:
            if not df_dispatch.empty:
                df_pend = df_dispatch[df_dispatch['pending_qty'] > 0]
                if not df_pend.empty:
                    st.dataframe(df_pend, use_container_width=True)
                    pend_buf = io.BytesIO()
                    df_pend.to_excel(pend_buf, index=False, sheet_name="Master Pending")
                    pend_buf.seek(0)
                    st.download_button("📥 Export Master Pending File (.xlsx)", pend_buf.getvalue(), f"Master_Pending_{get_ist_now().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_pending")
                else:
                    st.success("🟢 No pending items.")
            else:
                st.info("No records.")
        with tab_m9:
            st.markdown("#### 🛠️ Universal Schema Modifier (Drop Column Tool)")
            sel_table_drop = st.selectbox("Select Table", ['unique_routes_master', 'unmapped_missing_dr_ledger', 'output_files_ledger', 'input_output_traceability', 'discrepancy_audit_ledger', 'dispatch_planning_ledger', 'inventory_stock_table'])
            conn_schema = sqlite3.connect("sales_history.db")
            cur_schema = conn_schema.cursor()
            cur_schema.execute(f"PRAGMA table_info({sel_table_drop})")
            table_cols = [col[1] for col in cur_schema.fetchall()]
            conn_schema.close()
            
            col_to_drop = st.selectbox("Select Column to Drop", table_cols)
            if st.button("🗑️ Drop Selected Column"):
                try:
                    conn_d = sqlite3.connect("sales_history.db")
                    cur_d = conn_d.cursor()
                    remaining_cols = [c for c in table_cols if c != col_to_drop and c != 'id']
                    col_str = ", ".join(remaining_cols)
                    
                    cur_d.execute(f"CREATE TABLE {sel_table_drop}_temp AS SELECT id, {col_str} FROM {sel_table_drop}")
                    cur_d.execute(f"DROP TABLE {sel_table_drop}")
                    cur_d.execute(f"ALTER TABLE {sel_table_drop}_temp RENAME TO {sel_table_drop}")
                    conn_d.commit()
                    conn_d.close()
                    st.success(f"✅ Column '{col_to_drop}' dropped successfully from {sel_table_drop}!")
                    st.rerun()
                except Exception as ex:
                    st.error(f"Error dropping column: {str(ex)}")

    except Exception as e:
        st.error(f"Error loading databases: {str(e)}")

# Tables and Logs with Search & Filter Feature
if st.session_state.comparison_summary:
    st.markdown("---")
    st.markdown("### 📋 Agency-wise Material & Input Comparison")
    search_query = st.text_input("🔍 Search Table (Filter by Agency, DR Code, or FG Code)", "", key="table_search")
    combined_df = pd.concat(st.session_state.comparison_summary, ignore_index=True)
    summary_table = combined_df.groupby(["Agency", "DR Code", "FG Code"], as_index=False).agg({"Input Qty": "sum", "Generated Qty": "sum"})
    summary_table["Difference"] = summary_table["Input Qty"] - summary_table["Generated Qty"]
    if search_query:
        q = search_query.lower()
        summary_table = summary_table[
            summary_table['Agency'].astype(str).str.lower().str.contains(q) |
            summary_table['DR Code'].astype(str).str.lower().str.contains(q) |
            summary_table['FG Code'].astype(str).str.lower().str.contains(q)
        ]
    st.dataframe(summary_table, use_container_width=True)

if st.session_state.skipped_rows_log:
    st.markdown("---")
    st.markdown("### ⚠️ Skipped / Invalid Rows Exception Log")
    df_skipped = pd.DataFrame(st.session_state.skipped_rows_log)
    skip_search = st.text_input("🔍 Search Skipped Log (Filter by File Name or Agency)", "", key="skip_search")
    if skip_search:
        sq = skip_search.lower()
        df_skipped = df_skipped[df_skipped['File Name'].astype(str).str.lower().str.contains(sq) | df_skipped['Agency Value'].astype(str).str.lower().str.contains(sq)]
    st.dataframe(df_skipped, use_container_width=True)

# --- Historical Trend Analysis View ---
st.markdown("---")
with st.expander("🕒 View Historical Trend Analysis (SQLite Database - IST)"):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_history = pd.read_sql("SELECT * FROM history_logs ORDER BY id DESC", conn)
        conn.close()
        if not df_history.empty:
            st.dataframe(df_history, use_container_width=True)
            st.markdown("##### Day-over-Day Total Quantity Trend")
            st.line_chart(df_history.set_index("timestamp")["total_qty"])
        else:
            st.info("No historical logs available yet.")
    except Exception as e:
        st.error(f"Error loading history: {str(e)}")
