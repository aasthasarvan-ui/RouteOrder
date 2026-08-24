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
        st.session_state.vehicle_no = st.text_input("Default Vehicle Number (Fallback)", value=st.session_state.vehicle_no)
    with col_v2:
        st.session_state.driver_mobile = st.text_input("Default Driver Mobile No (Fallback)", value=st.session_state.driver_mobile)

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
