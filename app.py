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
        st.subheader("Default Route Fallback")
        st.session_state.route = st.text_input("Route Input", value=st.session_state.route, label_visibility="collapsed")
    with col_set2:
        st.subheader("Direct Column Index Mapping")
        st.session_state.col_map = st.text_area("Col Map Input", value=st.session_state.col_map, label_visibility="collapsed", height=100)
    with col_set3:
        st.subheader("Agency & Column-wise FG Override")
        st.session_state.agency_override = st.text_area("Agency Override Input", value=st.session_state.agency_override, label_visibility="collapsed", height=100)

default_fg_code = st.session_state.fg_code
col_mapping_input = st.session_state.col_map
agency_fg_override = st.session_state.agency_override
default_fallback_route = st.session_state.route
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
# 5. CORE BATCH PROCESSING & VEHICLE OPTIMIZATION
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
                
                # Fetch available stock inventory from DB
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
                        if any(kw in agency_str.upper() for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                            continue

                        if not agency_str.isdigit():
                            continue
                        agency_val = int(agency_str)
                        
                        valid_row_quantities = []
                        row_total_qty = 0
                        for c, fg_code in valid_cols:
                            if c >= total_col:
                                continue
                            sku_qty = df_input.iloc[r, c]
                            if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                try:
                                    qty_val = float(sku_qty)
                                    if qty_val > 0:
                                        row_total_qty += qty_val
                                        file_input_qty += qty_val
                                        
                                        # Check Stock Inventory Shortfall
                                        current_fg_code = agency_col_override_map.get((agency_val, c), direct_col_mapping.get(c, default_fg_code))
                                        avail_stock = stock_inventory_dict.get(current_fg_code, 99999.0) # Default full stock if not uploaded
                                        
                                        if avail_stock >= qty_val:
                                            dispatched_qty = qty_val
                                            pending_qty = 0.0
                                            stock_inventory_dict[current_fg_code] = avail_stock - qty_val # Deduct stock
                                        else:
                                            dispatched_qty = avail_stock if avail_stock > 0 else 0.0
                                            pending_qty = qty_val - dispatched_qty
                                            stock_inventory_dict[current_fg_code] = 0.0

                                        valid_row_quantities.append((c, current_fg_code, qty_val, dispatched_qty, pending_qty))
                                except ValueError:
                                    pass

                        if not valid_row_quantities:
                            continue

                        # Vehicle Capacity Check
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

                        # Log to Dispatch Planning Ledger (Demand vs Dispatch vs Pending)
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

                        if has_dr_code:
                            agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_valid, valid_row, valid_order_num, clean_dr, "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_missing, missing_row, missing_order_num, f"NEW_CUST_{agency_val}", "Missing DR"

                        item_id = 10
                        for c, current_fg, q_in, q_disp, q_pend in valid_row_quantities:
                            if q_disp <= 0:
                                continue # Do not generate order line if fully pending due to vehicle capacity or stock

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

                            target_ws.cell(row=current_r, column=2, value=order_num)
                            target_ws.cell(row=current_r, column=3, value="OR")
                            target_ws.cell(row=current_r, column=4, value="SO20")
                            target_ws.cell(row=current_r, column=7, value=dr_to_use)
                            target_ws.cell(row=current_r, column=8, value=dr_to_use)
                            target_ws.cell(row=current_r, column=9, value=ref_number)
                            target_ws.cell(row=current_r, column=10, value=today_date)
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
                            valid_row, valid_order_num, total_valid_orders = current_r, valid_order_num + 1, total_valid_orders + 1
                        else:
                            missing_row, missing_order_num, total_missing_orders = current_r, missing_order_num + 1, total_missing_orders + 1

                    if valid_items_created > 0 or len(file_comparison_rows) > 0:
                        buf_valid = io.BytesIO()
                        wb_valid.save(buf_valid)
                        buf_valid.seek(0)
                        out_fname = safe_route_num + "_" + today_date + "_" + timestamp + "_Dispatch_Plan.xlsx"
                        st.session_state.processed_files.append({"name": short_filename + " (Dispatch Schedule)", "data": buf_valid.getvalue(), "filename": out_fname, "orders": len(file_comparison_rows)})
                        output_files_to_store.append((out_fname, "Dispatch Plan", buf_valid.getvalue(), batch_ts))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname, "Dispatch Plan", 1, batch_ts))

                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.executemany("INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)", db_records_to_insert)
                cursor.executemany("INSERT OR IGNORE INTO output_files_ledger (file_name, file_type, file_data, created_at) VALUES (?, ?, ?, ?)", [(f[0], f[1], f[2], f[3]) for f in output_files_to_store])
                cursor.executemany("INSERT INTO input_output_traceability (batch_timestamp, input_file_name, input_file_blob, total_input_qty, generated_output_file, output_type, version_no, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", traceability_records)
                cursor.executemany("INSERT INTO dispatch_planning_ledger (dispatch_date, route_no, vehicle_no, driver_mobile, agency_no, fg_code, demand_qty, dispatched_qty, pending_qty, status, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", dispatch_plan_records)
                
                cursor.execute("INSERT INTO history_logs (timestamp, files_count, total_qty, status) VALUES (?, ?, ?, ?)", (get_ist_now().strftime("%Y-%m-%d %H:%M:%S"), len(uploaded_inputs), total_input_qty, "Success"))
                conn.commit()
                conn.close()

                st.session_state.kpi_data = {"input_qty": total_input_qty, "gen_qty": total_gen_qty, "valid_count": total_valid_orders, "missing_count": total_missing_orders, "skipped_count": total_skipped_rows}
                st.success("✅ Vehicle Capacity Dispatch Plan & Master Ledger Updated Successfully!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya demand files upload karein!")

# ==========================================
# 6. DASHBOARD & EXPORTS
# ==========================================
if st.session_state.processed_files:
    st.markdown("---")
    st.markdown("### 📈 Dispatch Planning & Performance Summary")
    kpi = st.session_state.kpi_data
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Demand Qty", f"{kpi['input_qty']:,.0f}")
    col2.metric("Total Dispatched Qty", f"{kpi['gen_qty']:,.0f}")
    col3.metric("Vehicle Capacity Limit", f"{vehicle_capacity_limit:,.0f}")
    col4.metric("Status", "Optimized & Planned")

    st.markdown("---")
    st.markdown("### 📥 Download Print-Ready Dispatch Schedules")
    for item in st.session_state.processed_files:
        st.download_button(
            label=f"📥 Download {item['filename']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{item['filename']}"
        )

# ==========================================
# 7. MULTI-DATABASE MANAGEMENT & MASTER FILES
# ==========================================
st.markdown("---")
with st.expander("🗄️ View, Export & Manage All Databases (Monthly Master & Master Pending Files)"):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_master = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
        df_dispatch = pd.read_sql("SELECT * FROM dispatch_planning_ledger ORDER BY id DESC", conn)
        conn.close()
        
        tab1, tab2, tab3 = st.tabs(["📋 Route-Agency-DR Master", "📅 Monthly Dispatch Masterfile", "⏳ Master Pending File (Next Day)"])
        
        with tab1:
            if not df_master.empty:
                st.dataframe(df_master, use_container_width=True)
            else:
                st.info("No master records found.")

        with tab2:
            st.markdown("#### 📅 Monthly Dispatch Masterfile (All Dispatched Details)")
            if not df_dispatch.empty:
                st.dataframe(df_dispatch, use_container_width=True)
                
                # Monthly Master Export
                monthly_buf = io.BytesIO()
                df_dispatch.to_excel(monthly_buf, index=False, sheet_name="Monthly Dispatch Master")
                monthly_buf.seek(0)
                st.download_button(
                    label="📥 Export Monthly Dispatch Masterfile (.xlsx)",
                    data=monthly_buf.getvalue(),
                    file_name=f"Monthly_Dispatch_Masterfile_{get_ist_now().strftime('%Y-%m')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="export_monthly_master"
                )
            else:
                st.info("No dispatch records available.")

        with tab3:
            st.markdown("#### ⏳ Master Pending File (Shortfall / Pending Items for Next Day)")
            if not df_dispatch.empty:
                df_pending_only = df_dispatch[df_dispatch['pending_qty'] > 0]
                if not df_pending_only.empty:
                    st.dataframe(df_pending_only, use_container_width=True)
                    
                    pending_buf = io.BytesIO()
                    df_pending_only.to_excel(pending_buf, index=False, sheet_name="Master Pending")
                    pending_buf.seek(0)
                    st.download_button(
                        label="📥 Export Master Pending File (.xlsx)",
                        data=pending_buf.getvalue(),
                        file_name=f"Master_Pending_File_{get_ist_now().strftime('%Y-%m-%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="export_master_pending"
                    )
                else:
                    st.success("🟢 No pending items! All demand fully dispatched.")
            else:
                st.info("No data available.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
