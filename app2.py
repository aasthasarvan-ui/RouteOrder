import io
import json
import os
import re
import smtplib
import sqlite3
import urllib.parse
import zipfile
import datetime
import pytz
import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from email.message import EmailMessage
from fpdf import FPDF

# ==============================================================================
# 1. PAGE CONFIGURATION & GLOBAL SETUP
# ==============================================================================
st.set_page_config(
    page_title="Enterprise Logistics & Sales Automation Hub",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

IST = pytz.timezone('Asia/Kolkata')

def get_ist_now():
    return datetime.datetime.now(IST)

def get_ist_date_str():
    return get_ist_now().strftime("%Y-%m-%d")

def get_ist_time_str():
    return get_ist_now().strftime("%H:%M:%S")

def get_ist_timestamp_full():
    return get_ist_now().strftime("%Y-%m-%d %H:%M:%S")

DB_FILE = "enterprise_logistics_hub.db"

# ==============================================================================
# 2. EXTENSIVE THEME ENGINE PALETTES
# ==============================================================================
THEMES = {
    "💼 Classic Enterprise Navy": {
        "bg": "#f4f6f9",
        "text": "#1f2937",
        "card_bg": "#ffffff",
        "border": "#cbd5e1",
        "btn_bg": "#1e3a8a",
        "btn_hover": "#1d4ed8",
        "primary": "#2563eb",
        "input_bg": "#ffffff",
        "input_text": "#1f2937",
        "table_header_bg": "#1e3a8a",
        "table_header_text": "#ffffff"
    },
    "🌙 Modern Dark Logistics": {
        "bg": "#0b0f19",
        "text": "#f3f4f6",
        "card_bg": "#1f2937",
        "border": "#374151",
        "btn_bg": "#374151",
        "btn_hover": "#4b5563",
        "primary": "#3b82f6",
        "input_bg": "#111827",
        "input_text": "#f3f4f6",
        "table_header_bg": "#1f2937",
        "table_header_text": "#f3f4f6"
    },
    "📊 Corporate Slate": {
        "bg": "#eef2f5",
        "text": "#0f172a",
        "card_bg": "#ffffff",
        "border": "#94a3b8",
        "btn_bg": "#475569",
        "btn_hover": "#334155",
        "primary": "#0284c7",
        "input_bg": "#ffffff",
        "input_text": "#0f172a",
        "table_header_bg": "#475569",
        "table_header_text": "#ffffff"
    },
    "🌲 Supply Chain Emerald": {
        "bg": "#f0fdf4",
        "text": "#14532d",
        "card_bg": "#dcfce7",
        "border": "#86efac",
        "btn_bg": "#16a34a",
        "btn_hover": "#15803d",
        "primary": "#22c55e",
        "input_bg": "#ffffff",
        "input_text": "#14532d",
        "table_header_bg": "#16a34a",
        "table_header_text": "#ffffff"
    },
    "🍇 Executive Burgundy": {
        "bg": "#fdf2f8",
        "text": "#500724",
        "card_bg": "#fce7f3",
        "border": "#f472b6",
        "btn_bg": "#db2777",
        "btn_hover": "#be185d",
        "primary": "#ec4899",
        "input_bg": "#ffffff",
        "input_text": "#500724",
        "table_header_bg": "#db2777",
        "table_header_text": "#ffffff"
    },
    "⚡ Cyber Teal": {
        "bg": "#f0fdfa",
        "text": "#042f2e",
        "card_bg": "#ccfbf1",
        "border": "#5eead4",
        "btn_bg": "#0d9488",
        "btn_hover": "#0f766e",
        "primary": "#14b8a6",
        "input_bg": "#ffffff",
        "input_text": "#042f2e",
        "table_header_bg": "#0d9488",
        "table_header_text": "#ffffff"
    },
    "☀️ Clean Minimalist": {
        "bg": "#ffffff",
        "text": "#111827",
        "card_bg": "#f9fafb",
        "border": "#d1d5db",
        "btn_bg": "#0f172a",
        "btn_hover": "#1e293b",
        "primary": "#10b981",
        "input_bg": "#ffffff",
        "input_text": "#111827",
        "table_header_bg": "#0f172a",
        "table_header_text": "#ffffff"
    },
    "🪙 Titanium Charcoal": {
        "bg": "#18181b",
        "text": "#fafafa",
        "card_bg": "#27272a",
        "border": "#52525b",
        "btn_bg": "#52525b",
        "btn_hover": "#71717a",
        "primary": "#e4e4e7",
        "input_bg": "#09090b",
        "input_text": "#fafafa",
        "table_header_bg": "#27272a",
        "table_header_text": "#fafafa"
    }
}

# ==============================================================================
# 3. SESSION STATE CONFIGURATION & DEFAULTS
# ==============================================================================
SESSION_DEFAULTS = {
    "selected_theme": "💼 Classic Enterprise Navy",
    "fg_code": "FG500014",
    "col_map": "36:FG500014AJ\n37:FG500014AK",
    "agency_override": "101:36:FG500014N01\n101:37:FG500014N02",
    "route": "22",
    "email_user": st.secrets.get("email", {}).get("sender_email", "") if hasattr(st, "secrets") else "",
    "email_pass": st.secrets.get("email", {}).get("app_password", "") if hasattr(st, "secrets") else "",
    "recipient": st.secrets.get("email", {}).get("recipient_email", "") if hasattr(st, "secrets") else "",
    "whatsapp_num": "919876543210",
    "processed_files": [],
    "comparison_summary": [],
    "skipped_rows_log": [],
    "anomaly_logs": [],
    "unmapped_current_batch": [],
    "kpi_data": {"input_qty": 0, "gen_qty": 0, "valid_count": 0, "missing_count": 0, "skipped_count": 0},
    "active_tab_index": 0
}

for k, val in SESSION_DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = val

active_theme = THEMES.get(st.session_state.selected_theme, THEMES["💼 Classic Enterprise Navy"])

# Apply Global CSS Theme
st.markdown(
    f"""
    <style>
        .stApp {{
            background-color: {active_theme['bg']};
            color: {active_theme['text']};
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{
            color: {active_theme['text']} !important;
        }}
        input, textarea, select {{
            background-color: {active_theme['input_bg']} !important;
            color: {active_theme['input_text']} !important;
            border: 1px solid {active_theme['border']} !important;
            border-radius: 4px !important;
        }}
        .stButton>button {{
            width: 100%;
            height: 38px;
            background-color: {active_theme['btn_bg']} !important;
            color: #ffffff !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            border-radius: 4px;
            border: 1px solid {active_theme['border']};
            transition: all 0.2s ease-in-out;
        }}
        .stButton>button:hover {{
            background-color: {active_theme['btn_hover']} !important;
            color: #ffffff !important;
        }}
        button[kind="primary"] {{
            background-color: {active_theme['primary']} !important;
            color: #ffffff !important;
        }}
        div[data-testid="stExpander"] {{
            background-color: {active_theme['card_bg']};
            border: 1px solid {active_theme['border']};
            border-radius: 6px;
            margin-bottom: 10px;
        }}
        div[data-testid="stDataFrame"] {{
            border: 1px solid {active_theme['border']};
            border-radius: 6px;
            background-color: {active_theme['card_bg']};
        }}
        .metric-card {{
            background-color: {active_theme['card_bg']};
            border: 1px solid {active_theme['border']};
            border-radius: 6px;
            padding: 12px;
            text-align: center;
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================================================================
# 4. DATABASE INITIALIZATION & SCHEMA DEFINITIONS
# ==============================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_enterprise_databases():
    conn = get_db_connection()
    cur = conn.cursor()

    # 1. Master Route-Agency-DR Mapping
    cur.execute("""
        CREATE TABLE IF NOT EXISTS unique_routes_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            created_at TEXT,
            UNIQUE(route_no, agency_no, dr_code)
        )
    """)

    # 2. Uploaded Input File Archive
    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploaded_files_archive (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            upload_timestamp TEXT,
            total_records INTEGER,
            file_size_kb REAL,
            batch_status TEXT
        )
    """)

    # 3. Pending Orders Ledger (With duplicate route guard flag)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pending_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT,
            order_no TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            bags_qty REAL,
            weight_mt REAL,
            order_ref TEXT,
            status TEXT DEFAULT 'Pending',
            uploaded_at TEXT
        )
    """)

    # 4. Fleet Master
    cur.execute("""
        CREATE TABLE IF NOT EXISTS fleet_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_no TEXT UNIQUE,
            vehicle_type TEXT,
            capacity_bags INTEGER,
            capacity_mt REAL,
            transporter_name TEXT,
            driver_name TEXT,
            driver_phone TEXT,
            status TEXT DEFAULT 'Available'
        )
    """)

    # 5. Loading Bays Master
    cur.execute("""
        CREATE TABLE IF NOT EXISTS loading_bays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bay_no TEXT UNIQUE,
            bay_name TEXT,
            status TEXT DEFAULT 'Open'
        )
    """)

    # 6. Trip Loading Slips & Vehicle Route Plans
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trip_loading_slips (
            trip_id TEXT PRIMARY KEY,
            trip_date TEXT,
            route_no TEXT,
            vehicle_no TEXT,
            transporter_name TEXT,
            driver_name TEXT,
            driver_phone TEXT,
            loading_bay TEXT,
            total_bags REAL,
            total_weight_mt REAL,
            capacity_utilization_pct REAL,
            status TEXT,
            created_at TEXT
        )
    """)

    # 7. Trip Order Items Sequence Mapping
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trip_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id TEXT,
            order_no TEXT,
            agency_no TEXT,
            route_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            allocated_bags REAL,
            allocated_weight_mt REAL,
            delivery_seq INTEGER,
            status TEXT DEFAULT 'Assigned',
            FOREIGN KEY (trip_id) REFERENCES trip_loading_slips(trip_id) ON DELETE CASCADE
        )
    """)

    # 8. Daily Dispatch Sale Register
    cur.execute("""
        CREATE TABLE IF NOT EXISTS daily_dispatch_register (
            register_id INTEGER PRIMARY KEY AUTOINCREMENT,
            dispatch_date TEXT,
            trip_id TEXT,
            vehicle_no TEXT,
            transporter_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            order_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            dispatched_bags REAL,
            dispatched_weight_mt REAL,
            bay_no TEXT,
            dispatched_at TEXT
        )
    """)

    # 9. Unmapped Missing DR Ledger
    cur.execute("""
        CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            created_at TEXT,
            UNIQUE(route_no, agency_no)
        )
    """)

    # 10. Audit History Logs
    cur.execute("""
        CREATE TABLE IF NOT EXISTS history_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            files_count INTEGER,
            total_qty REAL,
            status TEXT,
            remarks TEXT
        )
    """)

    # Default Fleet Seed
    cur.execute("SELECT COUNT(*) FROM fleet_master")
    if cur.fetchone()[0] == 0:
        default_fleet = [
            ('PB-10-AZ-1122', '10 Wheeler Truck', 400, 20.0, 'National Logistics', 'Gurpreet Singh', '9876543210', 'Available'),
            ('PB-08-BX-4455', '12 Wheeler Multi-Axle', 600, 30.0, 'Speedway Cargo', 'Baljit Sharma', '9812345678', 'Available'),
            ('PB-29-CD-9900', 'Canter / Eicher', 200, 10.0, 'Punjab Roadlines', 'Ramesh Kumar', '9823456789', 'Available'),
            ('PB-11-GH-3321', '14 Wheeler Heavy', 800, 40.0, 'Apex Transporters', 'Jarnail Singh', '9834567890', 'Available')
        ]
        cur.executemany("""
            INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, default_fleet)

    # Default Bays Seed
    cur.execute("SELECT COUNT(*) FROM loading_bays")
    if cur.fetchone()[0] == 0:
        default_bays = [
            ('BAY-01', 'North Plant Main Gate', 'Open'),
            ('BAY-02', 'Storage Silo Bay 2', 'Open'),
            ('BAY-03', 'Express Bulk Bay 3', 'Open')
        ]
        cur.executemany("INSERT INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, ?)", default_bays)

    conn.commit()
    conn.close()

init_enterprise_databases()

# ==============================================================================
# 5. DATA EXPORT & PDF GENERATION UTILITIES
# ==============================================================================
def export_dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name="DataSheet") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def generate_loading_slip_pdf(trip_data: dict, items_df: pd.DataFrame) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Title Header
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 8, "ENTERPRISE DISPATCH & LOADING SLIP", ln=True, align="C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(190, 5, f"Official Gate Pass & Loading Manifest | Generated (IST): {get_ist_timestamp_full()}", ln=True, align="C")
    pdf.ln(4)
    
    # Metadata Box
    pdf.set_font("Arial", "B", 9)
    pdf.cell(95, 6, f"Trip ID: {trip_data.get('trip_id', '')}", border=1)
    pdf.cell(95, 6, f"Trip Date: {trip_data.get('trip_date', '')}", border=1, ln=True)
    pdf.cell(95, 6, f"Vehicle No: {trip_data.get('vehicle_no', '')}", border=1)
    pdf.cell(95, 6, f"Route No: {trip_data.get('route_no', '')}", border=1, ln=True)
    pdf.cell(95, 6, f"Transporter: {trip_data.get('transporter_name', '')}", border=1)
    pdf.cell(95, 6, f"Loading Bay: {trip_data.get('loading_bay', '')}", border=1, ln=True)
    pdf.cell(95, 6, f"Driver Name: {trip_data.get('driver_name', '')}", border=1)
    pdf.cell(95, 6, f"Driver Phone: {trip_data.get('driver_phone', '')}", border=1, ln=True)
    pdf.ln(4)
    
    # Items Table Header
    pdf.set_font("Arial", "B", 9)
    pdf.cell(12, 6, "Seq", border=1, align="C")
    pdf.cell(32, 6, "Agency No", border=1, align="C")
    pdf.cell(40, 6, "DR Code", border=1, align="C")
    pdf.cell(46, 6, "FG Code", border=1, align="C")
    pdf.cell(30, 6, "Bags Qty", border=1, align="C")
    pdf.cell(30, 6, "Weight (MT)", border=1, ln=True, align="C")
    
    # Items Content
    pdf.set_font("Arial", "", 8)
    for _, it in items_df.iterrows():
        pdf.cell(12, 5, str(it.get("delivery_seq", "")), border=1, align="C")
        pdf.cell(32, 5, str(it.get("agency_no", "")), border=1, align="C")
        pdf.cell(40, 5, str(it.get("dr_code", "")), border=1, align="C")
        pdf.cell(46, 5, str(it.get("fg_code", "")), border=1, align="C")
        pdf.cell(30, 5, f"{float(it.get('allocated_bags', 0)):,.0f}", border=1, align="R")
        pdf.cell(30, 5, f"{float(it.get('allocated_weight_mt', 0)):,.2f}", border=1, ln=True, align="R")
        
    # Totals Row
    pdf.set_font("Arial", "B", 9)
    pdf.cell(130, 6, "TOTAL LOAD MANIFEST", border=1, align="R")
    pdf.cell(30, 6, f"{float(trip_data.get('total_bags', 0)):,.0f}", border=1, align="R")
    pdf.cell(30, 6, f"{float(trip_data.get('total_weight_mt', 0)):,.2f}", border=1, ln=True, align="R")
    
    pdf.ln(12)
    pdf.set_font("Arial", "", 9)
    pdf.cell(63, 6, "Driver Sign: ________________", ln=False)
    pdf.cell(63, 6, "Security Gate In: ________________", ln=False)
    pdf.cell(63, 6, "Supervisor Sign: ________________", ln=True)
    
    return bytes(pdf.output())

# ==============================================================================
# 6. SIDEBAR NAVIGATION
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/delivery-truck.png", width=55)
    st.title("Logistics & Sales Hub")
    
    nav_option = st.radio(
        "Navigation Menu",
        [
            "⚡ Inbound Demand & Sales Order Engine",
            "🚚 Route Dispatch Trip Planner",
            "📋 Loading Slips & Active Trips",
            "📖 Daily Dispatch Sale Register",
            "⏳ Pending Orders Ledger",
            "🗄️ File Upload Archive",
            "📋 Unique Master DB & Unmapped Ledger",
            "🚛 Fleet & Loading Bay Master",
            "📊 Executive KPI & Analytics Dashboard"
        ]
    )
    
    st.markdown("---")
    st.subheader("🎨 Interface Theme")
    new_theme = st.selectbox("Select Color Palette", list(THEMES.keys()), index=list(THEMES.keys()).index(st.session_state.selected_theme))
    if new_theme != st.session_state.selected_theme:
        st.session_state.selected_theme = new_theme
        st.rerun()
        
    st.markdown("---")
    st.subheader("📱 WhatsApp Alert Integration")
    st.session_state.whatsapp_num = st.text_input("WhatsApp Receiver Number", value=st.session_state.whatsapp_num)

# ==============================================================================
# MODULE 1: INBOUND DEMAND & SALES ORDER AUTOMATION ENGINE
# ==============================================================================
if nav_option == "⚡ Inbound Demand & Sales Order Engine":
    st.title("⚡ Enterprise Inbound Demand & Sales Order Processing Engine")
    st.markdown("Upload multiple **Demand Excel Workbooks** to parse SKU allocations, lookup master DRs, prevent route duplication, generate Output template files, and populate the pending orders database.")
    
    with st.expander("⚙️ SKU Mapping & Route Override Configuration Panel", expanded=False):
        c_p1, c_p2, c_p3 = st.columns(3)
        with c_p1:
            st.session_state.fg_code = st.text_input("Default Fallback FG Code", value=st.session_state.fg_code)
            st.session_state.route = st.text_input("Default Route Fallback", value=st.session_state.route)
        with c_p2:
            st.session_state.col_map = st.text_area("Direct Column Index Mapping (Col:FG)", value=st.session_state.col_map, height=80)
        with c_p3:
            st.session_state.agency_override = st.text_area("Agency-Specific Overrides (Agency:Col:FG)", value=st.session_state.agency_override, height=80)

    # Parsing Config Maps
    direct_col_map = {}
    for line in st.session_state.col_map.split('\n'):
        if ':' in line:
            parts = line.split(':')
            if parts[0].strip().isdigit():
                direct_col_map[int(parts[0].strip())] = parts[1].strip()

    agency_col_map = {}
    for line in st.session_state.agency_override.split('\n'):
        parts = line.split(':')
        if len(parts) == 3 and parts[0].strip().isdigit() and parts[1].strip().isdigit():
            agency_col_map[(int(parts[0].strip()), int(parts[1].strip()))] = parts[2].strip()

    uploaded_files = st.file_uploader(
        "Upload Sales Demand Excel Files (*.xlsx, *.xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.info(f"📂 {len(uploaded_files)} file(s) ready for execution.")

    if uploaded_files and st.button("🚀 Process Batch Orders & Ingest to Pending Database", type="primary"):
        st.session_state.processed_files = []
        st.session_state.comparison_summary = []
        st.session_state.skipped_rows_log = []
        st.session_state.anomaly_logs = []
        st.session_state.unmapped_current_batch = []
        
        tot_in_qty = 0
        tot_gen_qty = 0
        valid_orders_cnt = 0
        missing_orders_cnt = 0
        skipped_rows_cnt = 0
        skipped_dup_routes = 0
        
        conn = get_db_connection()
        cur = conn.cursor()
        
        # Load existing pending routes to prevent duplicate route re-uploads
        cur.execute("SELECT DISTINCT route_no FROM pending_orders WHERE status='Pending'")
        active_pending_routes = set(str(r[0]).strip() for r in cur.fetchall())
        
        try:
            with open("Output.xlsx", "rb") as f:
                template_file_bytes = f.read()
        except FileNotFoundError:
            template_file_bytes = None
            
        pending_records_batch = []
        master_routes_batch = []
        unmapped_records_batch = []
        
        progress_bar = st.progress(0)
        total_files = len(uploaded_files)
        
        for idx, up_file in enumerate(uploaded_files):
            short_fname = up_file.name
            if short_fname.lower() == "output.xlsx":
                continue
                
            f_bytes = up_file.getvalue()
            try:
                df_raw = pd.read_excel(io.BytesIO(f_bytes), header=None)
            except Exception as ex:
                st.error(f"Error parsing '{short_fname}': {str(ex)}")
                continue
                
            # Detect FG Header Row & Col
            fg_row, fg_col = -1, -1
            for r in range(df_raw.shape[0]):
                for c in range(df_raw.shape[1]):
                    if "FG" in str(df_raw.iloc[r, c]).strip().upper():
                        fg_row, fg_col = r, c
                        break
                if fg_row != -1:
                    break
                    
            if fg_row == -1:
                st.warning(f"⚠️ 'FG' column header not found in '{short_fname}'. File skipped.")
                continue

            # Detect Total Column Boundary
            total_col_idx = df_raw.shape[1]
            for c_s in range(fg_col, df_raw.shape[1]):
                cell_txt = str(df_raw.iloc[fg_row, c_s]).strip().upper()
                if any(kw in cell_txt for kw in ["TOTAL", "SUM", "NET"]):
                    total_col_idx = c_s
                    break

            # Route Number Resolution Logic
            route_num = st.session_state.route if st.session_state.route != "" else "22"
            for r in range(fg_row):
                for c in range(min(total_col_idx, 30)):
                    val = str(df_raw.iloc[r, c]).strip()
                    if val != "" and 1 <= len(val) <= 3 and any(ch.isdigit() for ch in val):
                        route_num = val
                        break

            resolved_route = str(route_num).strip()
            
            # Prevent Route Duplicate Check
            if resolved_route in active_pending_routes:
                skipped_dup_routes += 1
                continue

            # Agency Col & DR Code Col Resolution
            agency_col_idx = fg_col - 1 if fg_col > 0 else 0
            dr_col_idx = -1
            for c_scan in range(fg_col - 1, -1, -1):
                sample_str = str(df_raw.iloc[fg_row + 1, c_scan] if fg_row + 1 < df_raw.shape[0] else "").strip().upper()
                if re.match(r'^DR\d+', sample_str):
                    dr_col_idx = c_scan
                    break

            valid_sku_cols = [(c, str(df_raw.iloc[fg_row, c]).strip()) for c in range(fg_col, total_col_idx)]
            
            file_records_count = 0
            today_str = get_ist_date_str()
            now_ts_str = get_ist_timestamp_full()

            for r in range(fg_row + 1, df_raw.shape[0]):
                agency_raw = df_raw.iloc[r, agency_col_idx]
                if pd.isna(agency_raw) or str(agency_raw).strip() in ["", "nan", "None"]:
                    continue
                    
                agency_str = str(agency_raw).replace('.0', '').strip()
                if not agency_str.isdigit():
                    skipped_rows_cnt += 1
                    st.session_state.skipped_rows_log.append({
                        "File Name": short_fname,
                        "Row Index": r + 1,
                        "Agency": str(agency_raw),
                        "Reason": "Non-numeric agency identifier"
                    })
                    continue
                    
                agency_val = int(agency_str)
                
                # DR Resolution
                resolved_dr = ""
                if dr_col_idx >= 0:
                    raw_dr = df_raw.iloc[r, dr_col_idx]
                    if pd.notna(raw_dr) and "DR" in str(raw_dr).upper():
                        resolved_dr = str(raw_dr).strip()
                        
                if not resolved_dr:
                    cur.execute("SELECT dr_code FROM unique_routes_master WHERE route_no=? AND agency_no=?", (resolved_route, str(agency_val)))
                    match = cur.fetchone()
                    if match:
                        resolved_dr = match[0]
                    else:
                        resolved_dr = f"NEW_CUST_{agency_val}"
                        unmapped_records_batch.append((short_fname, resolved_route, str(agency_val), resolved_dr, now_ts_str))
                        st.session_state.unmapped_current_batch.append({
                            "File Name": short_fname,
                            "Route": resolved_route,
                            "Agency": agency_val,
                            "Fallback DR": resolved_dr
                        })

                if resolved_dr.startswith("DR"):
                    master_routes_batch.append((short_fname, resolved_route, str(agency_val), resolved_dr, now_ts_str))

                # Process SKU Allocations
                row_qty_accum = 0
                for c_idx, fg_val in valid_sku_cols:
                    q_val = df_raw.iloc[r, c_idx]
                    if pd.notna(q_val):
                        try:
                            f_qty = float(q_val)
                            if f_qty > 0:
                                current_fg = fg_val if fg_val.startswith("FG") else direct_col_map.get(c_idx, st.session_state.fg_code)
                                if (agency_val, c_idx) in agency_col_map:
                                    current_fg = agency_col_map[(agency_val, c_idx)]
                                    
                                ref_id = f"RT-{resolved_route}-{agency_val}-{today_str}"
                                pending_records_batch.append((
                                    short_fname,
                                    f"ORD-{agency_val}-{r}",
                                    resolved_route,
                                    str(agency_val),
                                    resolved_dr,
                                    current_fg,
                                    f_qty,
                                    round(f_qty * 0.05, 2),
                                    ref_id,
                                    'Pending',
                                    now_ts_str
                                ))
                                tot_in_qty += f_qty
                                tot_gen_qty += f_qty
                                row_qty_accum += f_qty
                                file_records_count += 1
                                
                                if resolved_dr.startswith("DR"):
                                    valid_orders_cnt += 1
                                else:
                                    missing_orders_cnt += 1
                        except ValueError:
                            pass

                if row_qty_accum > 500:
                    st.session_state.anomaly_logs.append({
                        "File Name": short_fname,
                        "Route": resolved_route,
                        "Agency": agency_val,
                        "Allocated Qty": row_qty_accum,
                        "Alert": "⚠️ Demand Spike > 500 Bags"
                    })

            # Archive Log Insert
            cur.execute("""
                INSERT OR REPLACE INTO uploaded_files_archive (file_name, upload_timestamp, total_records, file_size_kb, batch_status)
                VALUES (?, ?, ?, ?, 'Processed')
            """, (short_fname, now_ts_str, file_records_count, round(len(f_bytes)/1024, 2)))
            
            progress_bar.progress((idx + 1) / total_files)

        # Bulk SQL Inserts
        if pending_records_batch:
            cur.executemany("""
                INSERT INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, pending_records_batch)

        if master_routes_batch:
            cur.executemany("""
                INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at)
                VALUES (?, ?, ?, ?, ?)
            """, master_routes_batch)

        if unmapped_records_batch:
            cur.executemany("""
                INSERT OR IGNORE INTO unmapped_missing_dr_ledger (file_name, route_no, agency_no, dr_code, created_at)
                VALUES (?, ?, ?, ?, ?)
            """, unmapped_records_batch)

        # Add Audit History Log
        cur.execute("""
            INSERT INTO history_logs (timestamp, files_count, total_qty, status, remarks)
            VALUES (?, ?, ?, 'Success', ?)
        """, (get_ist_timestamp_full(), len(uploaded_files), tot_in_qty, f"Added {len(pending_records_batch)} records"))

        conn.commit()
        conn.close()

        st.session_state.kpi_data = {
            "input_qty": tot_in_qty,
            "gen_qty": tot_gen_qty,
            "valid_count": valid_orders_cnt,
            "missing_count": missing_orders_cnt,
            "skipped_count": skipped_rows_cnt
        }

        st.success(f"🎉 Batch Processed Successfully! Ingested {len(pending_records_batch)} orders into Pending Orders Database.")
        if skipped_dup_routes > 0:
            st.warning(f"⚠️ {skipped_dup_routes} Route(s) were skipped because their active pending records already exist in the database.")

    # KPI Metrics Display
    if st.session_state.kpi_data["input_qty"] > 0:
        st.markdown("---")
        st.subheader("📊 Batch Summary & Performance Metrics")
        kpi = st.session_state.kpi_data
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Input Bags", f"{kpi['input_qty']:,.0f}")
        c2.metric("Total Tonnage", f"{kpi['input_qty']*0.05:,.2f} MT")
        c3.metric("Valid DR Orders", kpi['valid_count'])
        c4.metric("Fallback DR Orders", kpi['missing_count'])
        c5.metric("Skipped Rows", kpi['skipped_count'], delta_color="inverse")

    if st.session_state.anomaly_logs:
        st.markdown("---")
        st.subheader("🚨 Demand Spike Alerts (> 500 Bags)")
        st.dataframe(pd.DataFrame(st.session_state.anomaly_logs), use_container_width=True)

    if st.session_state.unmapped_current_batch:
        st.markdown("---")
        st.subheader("⚠️ Unmapped Fallback DR Allocations")
        st.dataframe(pd.DataFrame(st.session_state.unmapped_current_batch), use_container_width=True)

# ==============================================================================
# MODULE 2: ROUTE DISPATCH & TRIP PLANNER
# ==============================================================================
elif nav_option == "🚚 Route Dispatch Trip Planner":
    st.title("🚚 Route Dispatch Planning & Vehicle Allocation")
    st.markdown("Select pending route demand, allocate vehicles from fleet master, check load utilization, and generate official loading slips.")
    
    conn = get_db_connection()
    df_pending = pd.read_sql("SELECT * FROM pending_orders WHERE status='Pending'", conn)
    
    if df_pending.empty:
        st.info("ℹ️ No pending order demand found. Upload demand workbooks in Module 1.")
    else:
        route_summary = df_pending.groupby("route_no").agg({
            "agency_no": "nunique",
            "bags_qty": "sum",
            "weight_mt": "sum"
        }).reset_index().rename(columns={"agency_no": "Total Agencies", "bags_qty": "Total Bags", "weight_mt": "Total MT"})
        
        st.subheader("📦 Pending Demand Grouped by Route")
        st.dataframe(route_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🚛 Assign Vehicle & Loading Bay")
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            sel_route = st.selectbox("1. Select Route for Trip", route_summary["route_no"].tolist())
            route_orders = df_pending[df_pending["route_no"] == str(sel_route)]
            
            avail_fleet = pd.read_sql("SELECT * FROM fleet_master WHERE status='Available'", conn)
            avail_bays = pd.read_sql("SELECT * FROM loading_bays WHERE status='Open'", conn)
            
            fleet_options = [f"{r['vehicle_no']} | {r['vehicle_type']} (Cap: {r['capacity_bags']} Bags / {r['capacity_mt']} MT)" for _, r in avail_fleet.iterrows()]
            sel_vehicle_str = st.selectbox("2. Select Available Vehicle", fleet_options if fleet_options else ["No Vehicles Available"])
            sel_bay_str = st.selectbox("3. Select Loading Bay", [f"{r['bay_no']} - {r['bay_name']}" for _, r in avail_bays.iterrows()])
            
        with col_r2:
            st.markdown("##### 📋 Filter Agencies to Load:")
            agencies_in_route = route_orders["agency_no"].unique().tolist()
            selected_agencies = st.multiselect("Select Agencies for this Trip:", agencies_in_route, default=agencies_in_route)
            
            trip_demand_df = route_orders[route_orders["agency_no"].isin(selected_agencies)]
            trip_total_bags = trip_demand_df["bags_qty"].sum()
            trip_total_mt = trip_demand_df["weight_mt"].sum()
            
            if sel_vehicle_str != "No Vehicles Available":
                v_num = sel_vehicle_str.split(" | ")[0]
                v_info = avail_fleet[avail_fleet["vehicle_no"] == v_num].iloc[0]
                capacity_bags = v_info["capacity_bags"]
                utilization = (trip_total_bags / capacity_bags * 100) if capacity_bags > 0 else 0
                
                st.metric("Total Load Allocated", f"{trip_total_bags:,.0f} / {capacity_bags} Bags", f"{utilization:.1f}% Capacity Utilization")
                if utilization > 100:
                    st.error(f"🚨 **Overload Alert:** Exceeded truck capacity by {trip_total_bags - capacity_bags:,.0f} bags!")
                elif utilization < 70:
                    st.warning("⚠️ **Low Utilization:** Truck is under 70% capacity.")
                else:
                    st.success("🟢 **Optimal Vehicle Utilization!**")

        st.markdown("---")
        if st.button("🚀 Confirm Dispatch Trip & Generate Loading Slip", type="primary"):
            if sel_vehicle_str == "No Vehicles Available":
                st.error("❌ No available vehicle selected.")
            elif trip_demand_df.empty:
                st.error("❌ No agencies/orders selected for this trip.")
            else:
                cur = conn.cursor()
                now_ist = get_ist_now()
                trip_id = f"TRIP-{sel_route}-{now_ist.strftime('%Y%m%d%H%M%S')}"
                v_num = sel_vehicle_str.split(" | ")[0]
                v_info = avail_fleet[avail_fleet["vehicle_no"] == v_num].iloc[0]
                bay_code = sel_bay_str.split(" - ")[0]
                
                # Insert Trip Master
                cur.execute("""
                    INSERT INTO trip_loading_slips (
                        trip_id, trip_date, route_no, vehicle_no, transporter_name, driver_name, driver_phone,
                        loading_bay, total_bags, total_weight_mt, capacity_utilization_pct, status, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Planned', ?)
                """, (
                    trip_id,
                    now_ist.strftime("%Y-%m-%d"),
                    str(sel_route),
                    v_num,
                    v_info["transporter_name"],
                    v_info["driver_name"],
                    v_info["driver_phone"],
                    bay_code,
                    trip_total_bags,
                    trip_total_mt,
                    round((trip_total_bags / v_info["capacity_bags"] * 100), 2),
                    now_ist.strftime("%Y-%m-%d %H:%M:%S")
                ))
                
                # Insert Order Items & Update Status
                for seq_idx, (_, row) in enumerate(trip_demand_df.iterrows(), 1):
                    cur.execute("""
                        INSERT INTO trip_order_items (
                            trip_id, order_no, agency_no, route_no, dr_code, fg_code,
                            allocated_bags, allocated_weight_mt, delivery_seq, status
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Assigned')
                    """, (
                        trip_id, row["order_no"], row["agency_no"], row["route_no"],
                        row["dr_code"], row["fg_code"], row["bags_qty"], row["weight_mt"], seq_idx
                    ))
                    cur.execute("UPDATE pending_orders SET status='Assigned' WHERE id=?", (row["id"],))
                    
                cur.execute("UPDATE fleet_master SET status='Assigned to Trip' WHERE vehicle_no=?", (v_num,))
                conn.commit()
                
                st.success(f"🎉 Trip '{trip_id}' Created Successfully! Vehicle assigned.")
                st.rerun()
    conn.close()

# ==============================================================================
# MODULE 3: LOADING SLIPS & ACTIVE TRIPS
# ==============================================================================
elif nav_option == "📋 Loading Slips & Active Trips":
    st.title("📋 Trip Slips & Active Vehicle Dispatches")
    
    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips ORDER BY created_at DESC", conn)
    
    # Live Search Bar
    search_trip = st.text_input("🔍 Search Trips (Trip ID, Vehicle, Route, Transporter, Status):", "")
    if search_trip:
        df_trips = df_trips[df_trips.apply(lambda r: r.astype(str).str.contains(search_trip, case=False).any(), axis=1)]
        
    st.dataframe(df_trips, use_container_width=True)
    
    # Multi-Delete & Excel Export
    col_t1, col_t2 = st.columns([1, 2])
    with col_t1:
        if not df_trips.empty:
            st.download_button(
                "📥 Export Trips to Excel",
                export_dataframe_to_excel_bytes(df_trips, "Trips"),
                "Trip_Loading_Slips.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col_t2:
        with st.expander("🗑️ Multi-Select Delete Trips"):
            del_trip_ids = st.multiselect("Select Trip IDs to Delete:", df_trips["trip_id"].tolist() if not df_trips.empty else [])
            if st.button("Delete Selected Trips"):
                cur = conn.cursor()
                for tid in del_trip_ids:
                    v_row = cur.execute("SELECT vehicle_no FROM trip_loading_slips WHERE trip_id=?", (tid,)).fetchone()
                    if v_row:
                        cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (v_row[0],))
                    cur.execute("DELETE FROM trip_order_items WHERE trip_id=?", (tid,))
                    cur.execute("DELETE FROM trip_loading_slips WHERE trip_id=?", (tid,))
                conn.commit()
                st.success("Selected trips deleted and vehicles released.")
                st.rerun()

    if not df_trips.empty:
        st.markdown("---")
        st.subheader("📄 Inspect Manifest, Generate PDF & Gate Out")
        
        sel_trip_id = st.selectbox("Select Trip ID to Inspect:", df_trips["trip_id"].tolist())
        trip_row = df_trips[df_trips["trip_id"] == sel_trip_id].iloc[0]
        items_df = pd.read_sql("SELECT * FROM trip_order_items WHERE trip_id=? ORDER BY delivery_seq ASC", conn, params=(sel_trip_id,))
        
        st.dataframe(items_df, use_container_width=True)
        
        c_btn1, c_btn2, c_btn3 = st.columns(3)
        with c_btn1:
            pdf_data = generate_loading_slip_pdf(trip_row.to_dict(), items_df)
            st.download_button(
                "📄 Download Loading Slip (PDF)",
                pdf_data,
                f"Loading_Slip_{sel_trip_id}.pdf",
                "application/pdf"
            )
        with c_btn2:
            wa_text = f"🚛 *Enterprise Dispatch Notification*\n*Trip ID:* {trip_row['trip_id']}\n*Vehicle:* {trip_row['vehicle_no']}\n*Driver:* {trip_row['driver_name']} ({trip_row['driver_phone']})\n*Route:* {trip_row['route_no']}\n*Total Load:* {trip_row['total_bags']} Bags ({trip_row['total_weight_mt']} MT)\n*Bay:* {trip_row['loading_bay']}"
            wa_url = f"https://wa.me/{st.session_state.whatsapp_num}?text={urllib.parse.quote(wa_text)}"
            st.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer;">📱 Send WhatsApp Alert</button></a>', unsafe_allow_html=True)
        with c_btn3:
            if trip_row["status"] != "Dispatched":
                if st.button("🏁 Mark Dispatched & Move to Daily Register", type="primary"):
                    cur = conn.cursor()
                    for _, it in items_df.iterrows():
                        cur.execute("""
                            INSERT INTO daily_dispatch_register (
                                dispatch_date, trip_id, vehicle_no, transporter_name, route_no, agency_no,
                                order_no, dr_code, fg_code, dispatched_bags, dispatched_weight_mt, bay_no, dispatched_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            trip_row["trip_date"], trip_row["trip_id"], trip_row["vehicle_no"],
                            trip_row["transporter_name"], trip_row["route_no"], it["agency_no"],
                            it["order_no"], it["dr_code"], it["fg_code"], it["allocated_bags"],
                            it["allocated_weight_mt"], trip_row["loading_bay"], get_ist_timestamp_full()
                        ))
                    cur.execute("UPDATE trip_loading_slips SET status='Dispatched' WHERE trip_id=?", (sel_trip_id,))
                    cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_row["vehicle_no"],))
                    conn.commit()
                    st.success("✅ Vehicle marked Gate Out & logged to Daily Dispatch Sale Register!")
                    st.rerun()
            else:
                st.info("✅ Trip is already Dispatched.")
    conn.close()

# ==============================================================================
# MODULE 4: DAILY DISPATCH SALE REGISTER
# ==============================================================================
elif nav_option == "📖 Daily Dispatch Sale Register":
    st.title("📖 Daily Dispatch Sale Register Database")
    st.markdown("Official audit register of all dispatched trips and sales transactions.")
    
    conn = get_db_connection()
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register ORDER BY register_id DESC", conn)
    
    search_reg = st.text_input("🔍 Search Sale Register (Agency, Route, Order No, Vehicle, DR Code, FG Code):", "")
    if search_reg:
        df_reg = df_reg[df_reg.apply(lambda r: r.astype(str).str.contains(search_reg, case=False).any(), axis=1)]
        
    st.dataframe(df_reg, use_container_width=True)
    
    c_r1, c_r2 = st.columns([1, 2])
    with c_r1:
        if not df_reg.empty:
            st.download_button(
                "📥 Export Register to Excel",
                export_dataframe_to_excel_bytes(df_reg, "DailyRegister"),
                "Daily_Dispatch_Sale_Register.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with c_r2:
        with st.expander("🗑️ Multi-Select Delete Register Records"):
            del_reg_ids = st.multiselect("Select Register IDs to Delete:", df_reg["register_id"].tolist() if not df_reg.empty else [])
            if st.button("Delete Selected Register Records"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM daily_dispatch_register WHERE register_id=?", [(i,) for i in del_reg_ids])
                conn.commit()
                st.success("Records deleted successfully.")
                st.rerun()
    conn.close()

# ==============================================================================
# MODULE 5: PENDING ORDERS LEDGER
# ==============================================================================
elif nav_option == "⏳ Pending Orders Ledger":
    st.title("⏳ Pending Orders Database")
    st.markdown("All inbound customer demand rows waiting for vehicle trip assignment.")
    
    conn = get_db_connection()
    df_pending_all = pd.read_sql("SELECT * FROM pending_orders ORDER BY id DESC", conn)
    
    search_pend = st.text_input("🔍 Search Pending Orders (Order No, Route, Agency, FG Code, DR Code):", "")
    if search_pend:
        df_pending_all = df_pending_all[df_pending_all.apply(lambda r: r.astype(str).str.contains(search_pend, case=False).any(), axis=1)]
        
    st.dataframe(df_pending_all, use_container_width=True)
    
    c_p1, c_p2 = st.columns([1, 2])
    with c_p1:
        if not df_pending_all.empty:
            st.download_button(
                "📥 Export Pending Orders to Excel",
                export_dataframe_to_excel_bytes(df_pending_all, "PendingOrders"),
                "Pending_Orders_Ledger.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with c_p2:
        with st.expander("🗑️ Multi-Select Delete Pending Orders"):
            del_pend_ids = st.multiselect("Select Order IDs to Delete:", df_pending_all["id"].tolist() if not df_pending_all.empty else [])
            if st.button("Delete Selected Pending Orders"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM pending_orders WHERE id=?", [(i,) for i in del_pend_ids])
                conn.commit()
                st.success("Selected orders deleted.")
                st.rerun()
    conn.close()

# ==============================================================================
# MODULE 6: FILE UPLOAD ARCHIVE
# ==============================================================================
elif nav_option == "🗄️ File Upload Archive":
    st.title("🗄️ Uploaded Input File Archive")
    st.markdown("History and batch metadata of all uploaded inbound demand files.")
    
    conn = get_db_connection()
    df_archive = pd.read_sql("SELECT * FROM uploaded_files_archive ORDER BY id DESC", conn)
    
    search_arch = st.text_input("🔍 Search Archive Files:", "")
    if search_arch:
        df_archive = df_archive[df_archive.apply(lambda r: r.astype(str).str.contains(search_arch, case=False).any(), axis=1)]
        
    st.dataframe(df_archive, use_container_width=True)
    
    c_a1, c_a2 = st.columns([1, 2])
    with c_a1:
        if not df_archive.empty:
            st.download_button(
                "📥 Export Archive Log to Excel",
                export_dataframe_to_excel_bytes(df_archive, "ArchiveLogs"),
                "Uploaded_Files_Archive.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with c_a2:
        with st.expander("🗑️ Multi-Select Delete Archive Logs"):
            del_arch_ids = st.multiselect("Select Archive IDs to Delete:", df_archive["id"].tolist() if not df_archive.empty else [])
            if st.button("Delete Selected Archive Logs"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM uploaded_files_archive WHERE id=?", [(i,) for i in del_arch_ids])
                conn.commit()
                st.success("Selected archive logs deleted.")
                st.rerun()
    conn.close()

# ==============================================================================
# MODULE 7: UNIQUE MASTER DB & UNMAPPED LEDGER
# ==============================================================================
elif nav_option == "📋 Unique Master DB & Unmapped Ledger":
    st.title("📋 Master Route-Agency-DR Database & Unmapped Fallback Ledger")
    
    conn = get_db_connection()
    df_master = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
    df_unmapped = pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger ORDER BY id DESC", conn)
    
    tab_m1, tab_m2 = st.tabs(["📋 Unique Master Mapping DB", "🚨 Unmapped Missing DR Ledger"])
    
    with tab_m1:
        search_m = st.text_input("🔍 Search Master Database (Route, Agency, DR Code):", "", key="search_m")
        if search_m:
            df_master = df_master[df_master.apply(lambda r: r.astype(str).str.contains(search_m, case=False).any(), axis=1)]
            
        st.dataframe(df_master, use_container_width=True)
        
        c_m1, c_m2 = st.columns([1, 2])
        with c_m1:
            st.download_button(
                "📥 Export Master DB to Excel",
                export_dataframe_to_excel_bytes(df_master, "MasterRoutes"),
                "Unique_Routes_Master.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_m2:
            with st.expander("🗑️ Multi-Select Delete Master Records"):
                del_m_ids = st.multiselect("Select Master IDs to Delete:", df_master["id"].tolist() if not df_master.empty else [])
                if st.button("Delete Selected Master Records"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM unique_routes_master WHERE id=?", [(i,) for i in del_m_ids])
                    conn.commit()
                    st.success("Selected master records deleted.")
                    st.rerun()

        with st.expander("➕ Add Single Record to Master DB"):
            c_add1, c_add2, c_add3 = st.columns(3)
            with c_add1:
                man_r = st.text_input("Route No", "10")
            with c_add2:
                man_a = st.text_input("Agency No", "")
            with c_add3:
                man_dr = st.text_input("DR Code (e.g. DR50012)", "")
            if st.button("Save Master Record"):
                if man_r and man_a and man_dr:
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT OR REPLACE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at)
                        VALUES ('Manual_Entry', ?, ?, ?, ?)
                    """, (man_r, man_a, man_dr, get_ist_timestamp_full()))
                    conn.commit()
                    st.success("Record saved to Master DB!")
                    st.rerun()

    with tab_m2:
        search_u = st.text_input("🔍 Search Unmapped Ledger:", "", key="search_u")
        if search_u:
            df_unmapped = df_unmapped[df_unmapped.apply(lambda r: r.astype(str).str.contains(search_u, case=False).any(), axis=1)]
            
        st.dataframe(df_unmapped, use_container_width=True)
        
        c_u1, c_u2 = st.columns([1, 2])
        with c_u1:
            st.download_button(
                "📥 Export Unmapped Ledger to Excel",
                export_dataframe_to_excel_bytes(df_unmapped, "UnmappedDR"),
                "Unmapped_Missing_DR_Ledger.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_u2:
            with st.expander("🗑️ Multi-Select Delete Unmapped Records"):
                del_u_ids = st.multiselect("Select Unmapped IDs to Delete:", df_unmapped["id"].tolist() if not df_unmapped.empty else [])
                if st.button("Delete Selected Unmapped Records"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM unmapped_missing_dr_ledger WHERE id=?", [(i,) for i in del_u_ids])
                    conn.commit()
                    st.success("Unmapped records deleted.")
                    st.rerun()
    conn.close()

# ==============================================================================
# MODULE 8: FLEET & LOADING BAY MASTER
# ==============================================================================
elif nav_option == "🚛 Fleet & Loading Bay Master":
    st.title("🚛 Fleet Master & Loading Bay Configurations")
    
    conn = get_db_connection()
    df_fleet = pd.read_sql("SELECT * FROM fleet_master", conn)
    df_bays = pd.read_sql("SELECT * FROM loading_bays", conn)
    
    tab_f1, tab_f2 = st.tabs(["🚛 Transporter Fleet Master", "🏭 Loading Bays Setup"])
    
    with tab_f1:
        search_f = st.text_input("🔍 Search Fleet (Vehicle No, Transporter, Driver):", "", key="search_f")
        if search_f:
            df_fleet = df_fleet[df_fleet.apply(lambda r: r.astype(str).str.contains(search_f, case=False).any(), axis=1)]
            
        st.dataframe(df_fleet, use_container_width=True)
        
        c_f1, c_f2 = st.columns([1, 2])
        with c_f1:
            st.download_button(
                "📥 Export Fleet to Excel",
                export_dataframe_to_excel_bytes(df_fleet, "FleetMaster"),
                "Fleet_Master.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_f2:
            with st.expander("🗑️ Multi-Select Delete Vehicles"):
                del_fleet_ids = st.multiselect("Select Vehicle IDs to Delete:", df_fleet["id"].tolist() if not df_fleet.empty else [])
                if st.button("Delete Selected Vehicles"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM fleet_master WHERE id=?", [(i,) for i in del_fleet_ids])
                    conn.commit()
                    st.success("Vehicles deleted.")
                    st.rerun()
                    
        with st.expander("➕ Add New Vehicle to Fleet"):
            vf1, vf2, vf3 = st.columns(3)
            with vf1:
                v_num = st.text_input("Vehicle No (e.g. PB-10-AZ-9988)")
                v_type = st.selectbox("Vehicle Type", ["10 Wheeler Truck", "12 Wheeler Multi-Axle", "14 Wheeler Heavy", "Canter / Eicher", "Mini Truck"])
            with vf2:
                v_cap_bags = st.number_input("Capacity (Bags)", min_value=50, max_value=2000, value=500, step=50)
                v_cap_mt = st.number_input("Capacity (Metric Tons)", min_value=1.0, max_value=100.0, value=25.0, step=1.0)
            with vf3:
                v_trans = st.text_input("Transporter Name", "National Logistics")
                v_driver = st.text_input("Driver Name", "Sukhdev Singh")
                v_phone = st.text_input("Driver Phone", "9876543210")
            if st.button("Save Vehicle"):
                if v_num:
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT OR REPLACE INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 'Available')
                    """, (v_num, v_type, v_cap_bags, v_cap_mt, v_trans, v_driver, v_phone))
                    conn.commit()
                    st.success(f"Vehicle {v_num} saved successfully!")
                    st.rerun()

    with tab_f2:
        search_b = st.text_input("🔍 Search Bays:", "", key="search_b")
        if search_b:
            df_bays = df_bays[df_bays.apply(lambda r: r.astype(str).str.contains(search_b, case=False).any(), axis=1)]
            
        st.dataframe(df_bays, use_container_width=True)
        
        c_b1, c_b2 = st.columns([1, 2])
        with c_b1:
            st.download_button(
                "📥 Export Bays to Excel",
                export_dataframe_to_excel_bytes(df_bays, "LoadingBays"),
                "Loading_Bays.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c_b2:
            with st.expander("🗑️ Multi-Select Delete Bays"):
                del_bay_ids = st.multiselect("Select Bay IDs to Delete:", df_bays["id"].tolist() if not df_bays.empty else [])
                if st.button("Delete Selected Bays"):
                    cur = conn.cursor()
                    cur.executemany("DELETE FROM loading_bays WHERE id=?", [(i,) for i in del_bay_ids])
                    conn.commit()
                    st.success("Bays deleted.")
                    st.rerun()
                    
        with st.expander("➕ Add Loading Bay"):
            b1, b2 = st.columns(2)
            with b1:
                bay_num = st.text_input("Bay No (e.g. BAY-04)")
            with b2:
                bay_name = st.text_input("Bay Location Name", "South Silo Discharge Bay 4")
            if st.button("Save Bay"):
                if bay_num:
                    cur = conn.cursor()
                    cur.execute("INSERT OR REPLACE INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, 'Open')", (bay_num, bay_name))
                    conn.commit()
                    st.success("Bay saved successfully!")
                    st.rerun()
    conn.close()

# ==============================================================================
# MODULE 9: EXECUTIVE KPI & ANALYTICS DASHBOARD
# ==============================================================================
elif nav_option == "📊 Executive KPI & Analytics Dashboard":
    st.title("📊 Supply Chain & Dispatch Analytics Dashboard")
    
    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips", conn)
    df_pending = pd.read_sql("SELECT * FROM pending_orders", conn)
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register", conn)
    conn.close()
    
    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    tot_trips = len(df_trips)
    tot_dispatched_bags = df_reg["dispatched_bags"].sum() if not df_reg.empty else 0
    avg_util = df_trips["capacity_utilization_pct"].mean() if not df_trips.empty else 0
    active_pending_bags = df_pending[df_pending["status"] == "Pending"]["bags_qty"].sum() if not df_pending.empty else 0
    
    col_k1.metric("Total Trips Formulated", tot_trips)
    col_k2.metric("Dispatched Bags", f"{tot_dispatched_bags:,.0f}")
    col_k3.metric("Avg Fleet Utilization", f"{avg_util:.1f}%")
    col_k4.metric("Pending Orders Load", f"{active_pending_bags:,.0f} Bags")
    
    st.markdown("---")
    c_ch1, c_ch2 = st.columns(2)
    with c_ch1:
        st.markdown("##### 🚛 Route-wise Dispatched Load (MT)")
        if not df_reg.empty:
            st.bar_chart(df_reg.groupby("route_no")["dispatched_weight_mt"].sum())
        else:
            st.info("No dispatched register records available.")
    with c_ch2:
        st.markdown("##### 📦 Transporter-wise Dispatched Bags")
        if not df_reg.empty:
            st.bar_chart(df_reg.groupby("transporter_name")["dispatched_bags"].sum())
        else:
            st.info("No transporter volume data available.")
