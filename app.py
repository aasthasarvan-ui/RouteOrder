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

# Page Configuration & Styling
st.set_page_config(
    page_title="Enterprise Sales Order Automation Hub", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 8 Professional Enterprise Themes with Unique Icons & Color Palettes
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

def verify_core_integrity():
    try:
        conn = sqlite3.connect("sales_history.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        existing_tables = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        required_tables = ['history_logs', 'unique_routes_master', 'output_files_ledger', 'unmapped_missing_dr_ledger', 'input_output_traceability', 'discrepancy_audit_ledger', 'dispatch_planning_ledger']
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
    cursor.execute("CREATE TABLE IF NOT EXISTS dispatch_planning_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, dispatch_date TEXT, route_no TEXT, vehicle_no TEXT, driver_mobile TEXT, agency_no TEXT, demand_qty REAL, dispatched_qty REAL, pending_qty REAL, status TEXT, created_at TEXT)")
    
    cursor.execute("PRAGMA table_info(unique_routes_master)")
    columns = [col[1] for col in cursor.fetchall()]
    if "file_name" not in columns:
        cursor.execute("ALTER TABLE unique_routes_master ADD COLUMN file_name TEXT")
        
    conn.commit()
    conn.close()

init_db()

is_healthy, health_msg = verify_core_integrity()
if not is_healthy:
    st.error(f"❌ **System Integrity Warning:** {health_msg}")
    st.stop()

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

# --- TOP CONTROL PANEL & VEHICLE CAPACITY SETTINGS ---
with st.expander("⚙️ Enterprise Control Panel, Vehicle Capacity & System Settings", expanded=True):
    st.subheader("🎨 Theme Engine & Vehicle Planning Config")
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        def on_theme_change():
            st.session_state.selected_theme = st.session_state.theme_selectbox
        st.selectbox("Select Interface Theme", list(THEMES.keys()), key="theme_selectbox", index=list(THEMES.keys()).index(st.session_state.selected_theme), on_change=on_theme_change)
    with col_t2:
        st.session_state.vehicle_capacity = st.number_input("Default Vehicle Capacity (Bags/Cases)", min_value=10.0, max_value=1000.0, value=st.session_state.vehicle_capacity, step=5.0)

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.session_state.vehicle_no = st.text_input("Default Vehicle Number", value=st.session_state.vehicle_no)
    with col_v2:
        st.session_state.driver_mobile = st.text_input("Default Driver Mobile No", value=st.session_state.driver_mobile)

    st.markdown("---")
    col_set1, col_set2, col_set3 = st.columns(3)
    
    with col_set1:
        st.subheader("Default Fallback FG Code")
        st.session_state.fg_code = st.text_input("FG Code Input", value=st.session_state.fg_code, label_visibility="collapsed")
        
        st.subheader("Default Route Fallback")
        st.session_state.route = st.text_input("Route Input", value=st.session_state.route, label_visibility="collapsed")

    with col_set2:
        st.subheader("Direct Column Index Mapping")
        st.session_state.col_map = st.text_area("Col Map Input", value=st.session_state.col_map, label_visibility="collapsed", height=100)

    with col_set3:
        st.subheader("Agency & Column-wise FG Override")
        st.session_state.agency_override = st.text_area("Agency Override Input", value=st.session_state.agency_override, label_visibility="collapsed", height=100)

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

st.title(f"💼 Enterprise Sales Order Automation Hub ({st.session_state.selected_theme})")
st.markdown("Upload multiple **Inbound Demand Files** to process orders, optimize vehicle dispatch loads, and track pending deliverables.")
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
        
        with st.spinner("⚡ Processing orders, calculating vehicle loads & building dispatch schedule... Please wait."):
            try:
                try:
                    with open("Output.xlsx", "rb") as f:
                        template_bytes = f.read()
                except FileNotFoundError:
                    st.error("❌ 'Output.xlsx' template file repository mein nahi mili.")
                    st.stop()
                
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
                        continue

                    total_col = df_input.shape[1]
                    for cSearch in range(fg_col, df_input.shape[1]):
                        is_total = False
                        for scan_r in range(max(0, fg_row - 10), min(fg_row + 3, df_input.shape[0])):
                            cell_val = str(df_input.iloc[scan_r, cSearch]).strip().upper()
                            if any(kw in cell_val for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                                is_total = True
                                break
                        if is_total:
                            total_col = cSearch
                            break

                    route_num = default_fallback_route if default_fallback_route != "" else "22"
                    for r in range(fg_row):
                        for c in range(min(total_col, 30)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            if cell_val != "" and 1 <= len(cell_val) <= 3 and any(char.isdigit() for char in cell_val):
                                route_num = cell_val
                                break
                        if route_num != "22":
                            break

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

                    valid_cols = []
                    for c in range(fg_col, total_col):
                        fg_code = str(df_input.iloc[fg_row, c] if fg_row >= 0 else "").strip()
                        if any(kw in fg_code.upper() for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                            break
                        valid_cols.append((c, fg_code))

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
                        if not agency_str.isdigit():
                            continue
                        agency_val = int(agency_str)
                        
                        row_has_items = False
                        valid_row_quantities = []
                        row_total_qty = 0
                        for c, fg_code in valid_cols:
                            sku_qty = df_input.iloc[r, c]
                            if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                try:
                                    qty_val = float(sku_qty)
                                    if qty_val > 0:
                                        row_has_items = True
                                        row_total_qty += qty_val
                                        file_input_qty += qty_val
                                        valid_row_quantities.append((c, fg_code, qty_val))
                                except ValueError:
                                    pass

                        if not row_has_items:
                            continue

                        dispatched_val = row_total_qty
                        pending_val = 0.0
                        status_str = "Fully Dispatched"

                        if current_vehicle_load + row_total_qty > vehicle_capacity_limit:
                            allowed_qty = max(0.0, vehicle_capacity_limit - current_vehicle_load)
                            if allowed_qty > 0:
                                dispatched_val = allowed_qty
                                pending_val = row_total_qty - allowed_qty
                                status_str = "Partially Dispatched (Pending Next Day)"
                            else:
                                dispatched_val = 0.0
                                pending_val = row_total_qty
                                status_str = "Fully Pending (Vehicle Full)"

                        current_vehicle_load += dispatched_val
                        if current_vehicle_load >= vehicle_capacity_limit:
                            current_vehicle_load = dispatched_val

                        dispatch_plan_records.append((today_date, str(route_num), assigned_vehicle_no, default_driver_mobile, str(agency_val), row_total_qty, dispatched_val, pending_val, status_str, batch_ts))

                        has_dr_code = False
                        clean_dr = ""
                        if dr_code_col >= 0:
                            raw_dr = df_input.iloc[r, dr_code_col]
                            if pd.notna(raw_dr):
                                val_str = str(raw_dr).replace('.0', '').strip()
                                if "DR" in val_str.upper() and val_str.upper() != "0":
                                    has_dr_code = True
                                    clean_dr = val_str

                        if not has_dr_code:
                            conn_lookup = sqlite3.connect("sales_history.db")
                            cursor_lookup = conn_lookup.cursor()
                            cursor_lookup.execute("SELECT dr_code FROM unique_routes_master WHERE route_no = ? AND agency_no = ? AND dr_code LIKE 'DR%' LIMIT 1", (str(route_num), str(agency_val)))
                            db_match = cursor_lookup.fetchone()
                            conn_lookup.close()
                            if db_match:
                                has_dr_code = True
                                clean_dr = db_match[0]

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
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_valid, valid_row, valid_order_num, clean_dr, "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_missing, missing_row, missing_order_num, f"NEW_CUST_{agency_val}", "Missing DR"

                        item_id = 10
                        for c, fg_code, qty_val in valid_row_quantities:
                            current_fg = agency_col_override_map.get((agency_val, c), direct_col_mapping.get(c, default_fg_code))
                            
                            total_input_qty += dispatched_val
                            total_gen_qty += dispatched_val
                            
                            file_comparison_rows.append({
                                "File Name": short_filename,
                                "Status": file_category,
                                "Agency": agency_val,
                                "DR Code": dr_to_use,
                                "FG Code": current_fg,
                                "Input Qty": dispatched_val,
                                "Generated Qty": dispatched_val
                            })

                            target_ws.cell(row=current_r, column=2, value=order_num)
                            target_ws.cell(row=current_r, column=3, value="OR")
                            target_ws.cell(row=current_r, column=4, value="SO20")
                            target_ws.cell(row=current_r, column=7, value=dr_to_use)
                            target_ws.cell(row=current_r, column=8, value=dr_to_use)
                            target_ws.cell(row=current_r, column=9, value=ref_number)
                            target_ws.cell(row=current_r, column=10, value=today_date)
                            target_ws.cell(row=current_r, column=15, value=item_id)
                            target_ws.cell(row=current_r, column=16, value=current_fg)
                            target_ws.cell(row=current_r, column=19, value=dispatched_val)
                            target_ws.cell(row=current_r, column=20, value="Bag")
                            target_ws.cell(row=current_r, column=22, value=2100)
                            target_ws.cell(row=current_r, column=26, value=str(route_num))
                            target_ws.cell(row=current_r, column=27, value=agency_val)
                            
                            item_id += 10
                            current_r += 1

                        if has_dr_code:
                            valid_row, valid_order_num, total_valid_orders = current_r, valid_order_num + 1, total_valid_orders + 1
                        else:
                            missing_row, missing_order_num, total_missing_orders = current_r, missing_order_num + 1, total_missing_orders + 1

                    if valid_items_created > 0 or len(file_comparison_rows) > 0:
                        buf_valid = io.BytesIO()
                        wb_valid.save(buf_valid)
                        buf_valid.seek(0)
                        out_fname = safe_route_num + "_" + today_date + "_" + timestamp + "_Valid.xlsx"
                        st.session_state.processed_files.append({"name": short_filename + " (Dispatch Plan)", "data": buf_valid.getvalue(), "filename": out_fname, "orders": len(file_comparison_rows)})
                        output_files_to_store.append((out_fname, "Dispatch Plan", buf_valid.getvalue(), batch_ts))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname, "Dispatch Plan", 1, batch_ts))

                    if file_comparison_rows:
                        df_comp = pd.DataFrame(file_comparison_rows)
                        df_pivot = df_comp.pivot_table(index=["File Name", "Status", "Agency", "DR Code", "FG Code"], values=["Input Qty", "Generated Qty"], aggfunc="sum").reset_index()
                        df_pivot["Difference"] = df_pivot["Input Qty"] - df_pivot["Generated Qty"]
                        st.session_state.comparison_summary.append(df_pivot)

                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.executemany("INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)", db_records_to_insert)
                cursor.executemany("INSERT OR IGNORE INTO output_files_ledger (file_name, file_type, file_data, created_at) VALUES (?, ?, ?, ?)", [(f[0], f[1], f[2], f[3]) for f in output_files_to_store])
                cursor.executemany("INSERT INTO input_output_traceability (batch_timestamp, input_file_name, input_file_blob, total_input_qty, generated_output_file, output_type, version_no, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", traceability_records)
                cursor.executemany("INSERT INTO dispatch_planning_ledger (dispatch_date, route_no, vehicle_no, driver_mobile, agency_no, demand_qty, dispatched_qty, pending_qty, status, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", dispatch_plan_records)
                
                cursor.execute("INSERT INTO history_logs (timestamp, files_count, total_qty, status) VALUES (?, ?, ?, ?)", (get_ist_now().strftime("%Y-%m-%d %H:%M:%S"), len(uploaded_inputs), total_input_qty, "Success"))
                conn.commit()
                conn.close()

                st.session_state.kpi_data = {"input_qty": total_input_qty, "gen_qty": total_gen_qty, "valid_count": total_valid_orders, "missing_count": total_missing_orders, "skipped_count": total_skipped_rows}
                st.success("✅ Vehicle Capacity Dispatch Plan & Master Ledger Updated Successfully!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya demand files upload karein!")

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

    st.markdown("---")
    st.markdown("### 📥 Bulk Download & Notifications")
    
    with st.expander("✉️ Advanced Email Dispatch Options (Custom Subject & Note)"):
        email_subject_custom = st.text_input("Custom Email Subject Line", f"🚀 Sales Orders Batch Execution Report (IST) - {get_ist_now().strftime('%Y-%m-%d')}")
        email_notes_custom = st.text_area("Custom Remarks / Notes to Include in Email Body", "All routes verified and processed successfully.")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])
    
    col_zip, col_pdf, col_summary, col_json, col_print, col_email, col_wa = st.columns(7)
    
    with col_zip:
        st.download_button(label="📦 ZIP", data=zip_buffer.getvalue(), file_name=f"Batch_Orders_{get_ist_now().strftime('%Y-%m-%d')}.zip", mime="application/zip", key="zip_download")
        
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
            for m_desc, m_val in [("Total Input Quantity", f"{kpi['input_qty']:,.0f}"), ("Total Generated Quantity", f"{kpi['gen_qty']:,.0f}"), ("Valid Orders", str(kpi['valid_count']))]:
                pdf.cell(100, 8, m_desc, border=1)
                pdf.cell(90, 8, m_val, border=1, ln=True)
            pdf_bytes = bytes(pdf.output())
            st.download_button(label="📄 PDF", data=pdf_bytes, file_name=f"Sales_Invoice_{get_ist_now().strftime('%Y-%m-%d')}.pdf", mime="application/pdf", key="pdf_invoice_download")
        except Exception as e:
            st.error(f"PDF Error: {str(e)}")

    with col_summary:
        summary_txt = f"Summary Report - Total Qty: {kpi['input_qty']}"
        st.download_button(label="📄 TXT", data=summary_txt.encode('utf-8'), file_name="Summary.txt", mime="text/plain", key="summary_txt_download")
        
    with col_json:
        json_data = json.dumps(kpi, indent=4)
        st.download_button(label="💾 JSON", data=json_data.encode('utf-8'), file_name="Audit.json", mime="application/json", key="json_backup_download")
        
    with col_print:
        print_html = '<div style="width:100%;"><button onclick="parent.window.print()" style="width:100%; height:38px; background:#2563eb; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer;">🖨️ Print</button></div>'
        components.html(print_html, height=50)
        
    with col_email:
        if st.button("📧 Email"):
            if email_user and email_pass and recipient_email:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = email_subject_custom
                    msg['From'] = email_user
                    msg['To'] = recipient_email
                    msg.set_content("Sales Order Batch Execution Report Attached.")
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email_user, email_pass)
                        smtp.send_message(msg)
                    st.success("✅ Email dispatched!")
                except Exception as e:
                    st.error(f"Email failed: {str(e)}")

    with col_wa:
        if whatsapp_num:
            wa_link = f"https://wa.me/{whatsapp_num}?text=Sales%20Order%20Batch%20Ready"
            st.markdown(f'<a href="{wa_link}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp</button></a>', unsafe_allow_html=True)

# --- ALL DATABASES & DISPATCH PLAN MANAGEMENT PANEL ---
st.markdown("---")
with st.expander("🗄️ View, Export & Manage All Databases (Master, Dispatch Planning & Pending Tracker)"):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_master = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
        df_dispatch = pd.read_sql("SELECT * FROM dispatch_planning_ledger ORDER BY id DESC", conn)
        df_trace = pd.read_sql("SELECT * FROM input_output_traceability ORDER BY id DESC", conn)
        conn.close()
        
        tab_m1, tab_m2, tab_m3 = st.tabs(["📋 Route-Agency-DR Master", "🚚 Dispatch Planning & Pending Tracker", "🔗 Traceability Ledger"])
        
        with tab_m1:
            if not df_master.empty:
                st.dataframe(df_master, use_container_width=True)
            else:
                st.info("No master records yet.")

        with tab_m2:
            st.markdown("#### 🚚 Daily Vehicle Dispatch Plan & Shortfall / Pending Tracker")
            if not df_dispatch.empty:
                st.dataframe(df_dispatch, use_container_width=True)
                
                if st.button("🗑️ Wipe Dispatch Planning Ledger"):
                    conn = sqlite3.connect("sales_history.db")
                    cur = conn.cursor()
                    cur.execute("DELETE FROM dispatch_planning_ledger")
                    cur.execute("DELETE FROM sqlite_sequence WHERE name='dispatch_planning_ledger'")
                    conn.commit()
                    conn.close()
                    st.success("✅ Dispatch Ledger wiped successfully!")
                    st.rerun()
            else:
                st.info("No dispatch plans generated in this session yet.")

        with tab_m3:
            if not df_trace.empty:
                st.dataframe(df_trace, use_container_width=True)
            else:
                st.info("No traceability records yet.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
