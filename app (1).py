# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# ENHANCED DISPATCH PLANNER WITH LIGHTNING FAST MASTER DR AUTO-LOOKUP & CRUD
# ==============================================================================

import datetime
import io
import json
import os
import re
import smtplib
import sqlite3
import urllib.parse
import zipfile
from email.message import EmailMessage

from fpdf import FPDF
import openpyxl
import pandas as pd
import pytz
import streamlit as st
import streamlit.components.v1 as components

# ==============================================================================
# SECTION 1: GLOBAL CONFIGURATION, TIMEZONE (IST) & METADATA
# ==============================================================================

st.set_page_config(
    page_title="Enterprise Logistics & Sales Automation Hub",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

IST = pytz.timezone("Asia/Kolkata")

def get_ist_now():
    return datetime.datetime.now(IST)

def get_ist_date_str():
    return get_ist_now().strftime("%Y-%m-%d")

def get_ist_time_str():
    return get_ist_now().strftime("%H:%M:%S")

def get_ist_timestamp_full():
    return get_ist_now().strftime("%Y-%m-%d %H:%M:%S")

def get_ist_file_suffix():
    return get_ist_now().strftime("%H%M%S")

DB_NAME = "enterprise_logistics_sales_hub.db"

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

master_path = os.path.join(application_path, "Business_Partners_Master_Original_Keys_Restored.xlsx")

# ==============================================================================
# SECTION 2: 8 COMPLETE ENTERPRISE COLOR PALETTES
# ==============================================================================

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

# ==============================================================================
# SECTION 3: SESSION STATE INITIALIZATION
# ==============================================================================

GLOBAL_DEFAULTS = {
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
    "kpi_data": {
        "input_qty": 0.0,
        "gen_qty": 0.0,
        "valid_count": 0,
        "missing_count": 0,
        "skipped_count": 0
    }
}

for d_key, d_val in GLOBAL_DEFAULTS.items():
    if d_key not in st.session_state:
        st.session_state[d_key] = d_val

current_theme = THEMES.get(st.session_state.selected_theme, THEMES["💼 Classic Enterprise Navy"])

st.markdown(
    f"""
    <style>
        .stApp {{
            background-color: {current_theme['bg']};
            color: {current_theme['text']};
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{
            color: {current_theme['text']} !important;
        }}
        input, textarea, select {{
            background-color: {current_theme['input_bg']} !important;
            color: {current_theme['input_text']} !important;
            border: 1px solid {current_theme['border']} !important;
            border-radius: 4px !important;
        }}
        .stButton>button {{
            width: 100%;
            height: 38px;
            background-color: {current_theme['btn_bg']} !important;
            color: #ffffff !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            border-radius: 4px;
            border: 1px solid {current_theme['border']};
            transition: all 0.2s ease-in-out;
        }}
        .stButton>button:hover {{
            background-color: {current_theme['btn_hover']} !important;
            color: #ffffff !important;
        }}
        button[kind="primary"] {{
            background-color: {current_theme['primary']} !important;
            color: #ffffff !important;
        }}
        div[data-testid="stExpander"] {{
            background-color: {current_theme['card_bg']};
            border: 1px solid {current_theme['border']};
            border-radius: 6px;
            margin-bottom: 12px;
        }}
        div[data-testid="stDataFrame"] {{
            border: 1px solid {current_theme['border']};
            border-radius: 6px;
            background-color: {current_theme['card_bg']};
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================================================================
# SECTION 4: UNIFIED DATABASE ARCHITECTURE (ALL CORE TABLES)
# ==============================================================================

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_all_enterprise_databases():
    conn = get_db_connection()
    cur = conn.cursor()

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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS loading_bays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bay_no TEXT UNIQUE,
            bay_name TEXT,
            status TEXT DEFAULT 'Open'
        )
    """)
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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS partial_dispatch_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id TEXT,
            source_file TEXT,
            order_no TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            original_bags REAL,
            dispatched_bags REAL,
            remaining_bags REAL,
            status TEXT DEFAULT 'Partial Pending',
            created_at TEXT
        )
    """)
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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS output_files_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            file_type TEXT,
            file_data BLOB,
            created_at TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS input_output_traceability (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_timestamp TEXT,
            input_file_name TEXT,
            input_file_blob BLOB,
            total_input_qty REAL,
            generated_output_file TEXT,
            output_type TEXT,
            version_no INTEGER,
            created_at TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS discrepancy_audit_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_timestamp TEXT,
            file_name TEXT,
            agency_no TEXT,
            dr_code TEXT,
            fg_code TEXT,
            input_qty REAL,
            generated_qty REAL,
            difference REAL,
            logged_at TEXT
        )
    """)
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

init_all_enterprise_databases()

# ==============================================================================
# SECTION 5: UTILITY & EXPORT GENERATOR ENGINES
# ==============================================================================

def to_excel_download_bytes(df: pd.DataFrame, sheet_name="DataSheet") -> bytes:
    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output_stream.getvalue()

def build_pdf_loading_slip(trip_data: dict, items_df: pd.DataFrame) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 8, "ENTERPRISE DISPATCH & LOADING SLIP", ln=True, align="C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(190, 5, f"Official Gate Pass & Manifest | Generated (IST): {get_ist_timestamp_full()}", ln=True, align="C")
    pdf.ln(4)

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

    pdf.set_font("Arial", "B", 9)
    pdf.cell(12, 6, "Seq", border=1, align="C")
    pdf.cell(32, 6, "Agency No", border=1, align="C")
    pdf.cell(40, 6, "DR Code", border=1, align="C")
    pdf.cell(46, 6, "FG Code", border=1, align="C")
    pdf.cell(30, 6, "Bags Qty", border=1, align="C")
    pdf.cell(30, 6, "Weight (MT)", border=1, ln=True, align="C")

    pdf.set_font("Arial", "", 8)
    for _, it in items_df.iterrows():
        pdf.cell(12, 5, str(it.get("delivery_seq", "")), border=1, align="C")
        pdf.cell(32, 5, str(it.get("agency_no", "")), border=1, align="C")
        pdf.cell(40, 5, str(it.get("dr_code", "")), border=1, align="C")
        pdf.cell(46, 5, str(it.get("fg_code", "")), border=1, align="C")
        pdf.cell(30, 5, f"{float(it.get('allocated_bags', 0)):,.0f}", border=1, align="R")
        pdf.cell(30, 5, f"{float(it.get('allocated_weight_mt', 0)):,.2f}", border=1, ln=True, align="R")

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
# SECTION 6: SIDEBAR NAVIGATION ENGINE
# ==============================================================================

with st.sidebar:
    st.image("https://img.icons8.com/color/96/delivery-truck.png", width=55)
    st.title("Logistics Master Suite")

    main_menu = st.radio(
        "Navigation",
        [
            "⚡ Inbound Demand & Sales Order Engine",
            "🚚 Route Dispatch Trip Planner",
            "📦 Live Inventory Stock & ERP Demand Matcher",
            "📋 Loading Slips & Active Trips",
            "📖 Daily Dispatch Sale Register",
            "🧩 Partial / Split Dispatch Database",
            "⏳ Pending Orders Ledger",
            "🗄️ File Upload Archive",
            "📋 Master DB & Unmapped Ledger",
            "🚛 Fleet & Loading Bay Master",
            "🔍 Traceability & Audit Ledgers",
            "📊 Executive KPI & Visual Analytics",
            "🎯 Universal Date & Multi-Field Filter Center",
            "⚡ Smart Multi-Truck Load Optimizer Pro",     
            "🗄️ In-App Database Builder & Dynamic Linker"
        ]
    )
    st.markdown("---")
    st.subheader("🎨 Interface Theme Engine")
    theme_choice = st.selectbox(
        "Select Theme",
        list(THEMES.keys()),
        index=list(THEMES.keys()).index(st.session_state.selected_theme)
    )
    if theme_choice != st.session_state.selected_theme:
        st.session_state.selected_theme = theme_choice
        st.rerun()

    st.markdown("---")
    st.subheader("📱 Notification Config")
    st.session_state.whatsapp_num = st.text_input("WhatsApp Alert Mobile No", value=st.session_state.whatsapp_num)

# ==============================================================================
# MODULE 1: INBOUND DEMAND & SALES ORDER AUTOMATION ENGINE
# (WITH LIGHTNING-FAST RAM CACHED MASTER FILE DR AUTO-LOOKUP)
# ==============================================================================

if main_menu == "⚡ Inbound Demand & Sales Order Engine":
    st.title("⚡ Enterprise Inbound Demand & Sales Order Processing Engine")
    st.markdown("Upload multiple **Demand Workbooks** to execute **Master File DR Auto-Lookup** instantly, eliminate duplicate orders, generate structured `Output.xlsx` files, and sync pending demand.")

    with st.expander("⚙️ SKU, Route & Multi-Channel Dispatch Settings", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.session_state.fg_code = st.text_input("Default Fallback FG Code", value=st.session_state.fg_code)
            st.session_state.route = st.text_input("Default Route Fallback", value=st.session_state.route)
        with c2:
            st.session_state.col_map = st.text_area("Column Index Mapping (Col:FG)", value=st.session_state.col_map, height=80)
        with c3:
            st.session_state.agency_override = st.text_area("Agency SKU Overrides (Agency:Col:FG)", value=st.session_state.agency_override, height=80)

        st.markdown("---")
        c4, c5 = st.columns(2)
        with c4:
            st.session_state.email_user = st.text_input("Sender Gmail ID", value=st.session_state.email_user)
            st.session_state.email_pass = st.text_input("Gmail App Password", type="password", value=st.session_state.email_pass)
        with c5:
            st.session_state.recipient = st.text_input("Recipient Email", value=st.session_state.recipient)

    # Parse Mappings
    col_map_dict = {}
    for line in st.session_state.col_map.split("\n"):
        if ":" in line:
            parts = line.split(":")
            if parts[0].strip().isdigit():
                col_map_dict[int(parts[0].strip())] = parts[1].strip()

    agency_override_dict = {}
    for line in st.session_state.agency_override.split("\n"):
        parts = line.split(":")
        if len(parts) == 3 and parts[0].strip().isdigit() and parts[1].strip().isdigit():
            agency_override_dict[(int(parts[0].strip()), int(parts[1].strip()))] = parts[2].strip()

    uploaded_files = st.file_uploader(
        "Upload Inbound Demand Excel Workbooks",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files and st.button("🚀 Process Batch Orders & Ingest to Pending Database", type="primary"):
        st.session_state.processed_files = []
        st.session_state.comparison_summary = []
        st.session_state.skipped_rows_log = []
        st.session_state.anomaly_logs = []
        st.session_state.unmapped_current_batch = []

        total_in_qty = 0.0
        total_gen_qty = 0.0
        total_valid = 0
        total_missing = 0
        total_skipped = 0

        conn = get_db_connection()
        cur = conn.cursor()

        try:
            with open("Output.xlsx", "rb") as f:
                template_bytes = f.read()
        except FileNotFoundError:
            template_bytes = None

        # --- 🔥 LIGHTNING-FAST O(1) RAM CACHING FOR MASTER LOOKUPS ---
        sql_lookup_dict = {}
        df_sql = pd.read_sql("SELECT route_no, agency_no, dr_code FROM unique_routes_master", conn)
        for _, r in df_sql.iterrows():
            sql_lookup_dict[(str(r['route_no']).strip(), str(r['agency_no']).strip())] = str(r['dr_code']).strip()

        master_lookup_dict = {}
        if os.path.exists(master_path):
            try:
                xls = pd.ExcelFile(master_path)
                for s_name in xls.sheet_names:
                    if s_name.startswith("Route_"):
                        df_r = pd.read_excel(xls, sheet_name=s_name, header=2)
                        df_r.columns = df_r.columns.astype(str).str.strip()
                        if 'Route' in df_r.columns and 'Agency' in df_r.columns and 'DRCODE' in df_r.columns:
                            for _, mr in df_r.iterrows():
                                rt_val = str(mr['Route']).replace('.0', '').strip()
                                ag_val = str(mr['Agency']).replace('.0', '').strip()
                                dr_val = str(mr['DRCODE']).strip()
                                if dr_val and dr_val.upper() not in ["NAN", "NONE", ""]:
                                    master_lookup_dict[(rt_val, ag_val)] = dr_val
            except Exception as ex:
                st.warning(f"Master file warning: {ex}")
        # -------------------------------------------------------------

        pending_records_to_insert = []
        master_routes_to_insert = []
        unmapped_records_to_insert = []
        traceability_records = []

        batch_ts = get_ist_timestamp_full()
        today_date = get_ist_date_str()
        time_suffix = get_ist_file_suffix()

        for up_file in uploaded_files:
            short_fname = up_file.name
            if short_fname.lower() == "output.xlsx":
                continue

            cur.execute("SELECT id FROM uploaded_files_archive WHERE file_name=?", (short_fname,))
            if cur.fetchone():
                st.warning(f"⚠️ '{short_fname}' pehle se process ho chuki hai. Duplicate upload skip kiya gaya.")
                continue

            f_bytes = up_file.getvalue()

            try:
                wb_raw = openpyxl.load_workbook(io.BytesIO(f_bytes), data_only=True)
                ws_raw = wb_raw.active
                data_matrix = []
                for row_cells in ws_raw.iter_rows(values_only=True):
                    data_matrix.append(list(row_cells))
                df_input = pd.DataFrame(data_matrix)
            except Exception as e:
                try:
                    df_input = pd.read_excel(io.BytesIO(f_bytes), header=None)
                except Exception as ex:
                    st.error(f"Error reading '{short_fname}': {str(ex)}")
                    continue

            # 1. Detect FG Header Row & Column
            fg_row, fg_col = -1, -1
            for r in range(min(df_input.shape[0], 30)):
                for c in range(df_input.shape[1]):
                    cell_val = str(df_input.iloc[r, c]).strip().upper()
                    if "FG" in cell_val or "PRODUCT" in cell_val or "SKU" in cell_val:
                        fg_row, fg_col = r, c
                        break
                if fg_row != -1:
                    break

            if fg_row == -1:
                st.error(f"❌ '{short_fname}' me FG / SKU header row detect nahi hui.")
                continue

            # 2. Strict Total Column Cutoff
            total_col = df_input.shape[1]
            for c_s in range(fg_col, df_input.shape[1]):
                val_header = str(df_input.iloc[fg_row, c_s]).strip().upper()
                if any(kw in val_header for kw in ["TOTAL", "SUM", "NET", "TTL", "GR. TOTAL", "GRAND"]):
                    total_col = c_s
                    break
                is_sum_col = False
                for r_scan in range(max(0, fg_row - 5), min(fg_row + 5, df_input.shape[0])):
                    cell_txt = str(df_input.iloc[r_scan, c_s]).strip().upper()
                    if any(kw in cell_txt for kw in ["TOTAL", "SUM", "GRAND TOTAL"]):
                        is_sum_col = True
                        break
                if is_sum_col:
                    total_col = c_s
                    break

            # 3. Route Number Detection
            route_num = st.session_state.route if st.session_state.route != "" else "22"
            for r in range(fg_row):
                for c in range(min(total_col, 20)):
                    val = str(df_input.iloc[r, c]).replace('.0', '').strip()
                    if val != "" and 1 <= len(val) <= 4 and any(ch.isdigit() for ch in val):
                        if not any(prefix in val.upper() for prefix in ["DR", "FG", "OR", "SO", "RT-", "TOTAL"]):
                            route_num = val
                            break

            resolved_route = "".join(ch for ch in str(route_num) if ch.isalnum() or ch in ('-', '_'))

            # 4. Agency & DR Columns
            agency_col = -1
            for c_s in range(fg_col - 1, -1, -1):
                col_samples = df_input.iloc[fg_row + 1: fg_row + 15, c_s].dropna().astype(str).str.replace(r'\.0$', '', regex=True)
                if col_samples.str.isdigit().sum() >= 2:
                    agency_col = c_s
                    break
            if agency_col == -1:
                agency_col = fg_col - 1 if fg_col > 0 else 0

            dr_code_col = -1
            for c_s in range(fg_col - 1, -1, -1):
                col_samples = df_input.iloc[fg_row + 1: fg_row + 15, c_s].dropna().astype(str).str.upper()
                if col_samples.str.startswith("DR").sum() >= 1:
                    dr_code_col = c_s
                    break

            valid_cols = [(c, str(df_input.iloc[fg_row, c]).strip()) for c in range(fg_col, total_col)]

            # Output Workbooks
            wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes)) if template_bytes else openpyxl.Workbook()
            ws_valid = wb_valid["Order Data"] if "Order Data" in wb_valid.sheetnames else wb_valid.active
            wb_missing = openpyxl.load_workbook(io.BytesIO(template_bytes)) if template_bytes else openpyxl.Workbook()
            ws_missing = wb_missing["Order Data"] if "Order Data" in wb_missing.sheetnames else wb_missing.active

            valid_r_idx, missing_r_idx = 6, 6
            valid_order_no, missing_order_no = 1, 1
            valid_items_cnt, missing_items_cnt = 0, 0
            agency_counts_valid, agency_counts_missing = {}, {}
            file_input_qty = 0.0

            for r in range(fg_row + 1, df_input.shape[0]):
                row_raw_values = [str(val).strip().upper() for val in df_input.iloc[r, :min(total_col, 15)] if pd.notna(val)]
                if any(any(kw in cell_str for kw in ["TOTAL", "SUM", "GRAND TOTAL", "GR. TOTAL", "NET TOTAL", "TOTAL QTY"]) for cell_str in row_raw_values):
                    continue

                agency_raw = df_input.iloc[r, agency_col]
                if pd.isna(agency_raw) or str(agency_raw).strip() in ["", "nan", "None", "0"]:
                    continue

                agency_str = str(agency_raw).replace(".0", "").strip()
                if not agency_str.isdigit() or not (1 <= len(agency_str) <= 6):
                    total_skipped += 1
                    st.session_state.skipped_rows_log.append({
                        "File Name": short_fname, "Row": r + 1, "Agency": str(agency_raw), "Reason": "Non-numeric agency"
                    })
                    continue

                agency_val = int(agency_str)

                # --- INSTANT RAM LOOKUP FOR DR CODE ---
                clean_dr = ""
                if dr_code_col >= 0:
                    raw_dr = df_input.iloc[r, dr_code_col]
                    if pd.notna(raw_dr) and "DR" in str(raw_dr).upper():
                        clean_dr = str(raw_dr).strip()

                lookup_key = (resolved_route, str(agency_val))
                if not clean_dr and lookup_key in sql_lookup_dict:
                    clean_dr = sql_lookup_dict[lookup_key]
                if not clean_dr and lookup_key in master_lookup_dict:
                    clean_dr = master_lookup_dict[lookup_key]

                if not clean_dr:
                    clean_dr = f"NEW_CUST_{agency_val}"
                    unmapped_records_to_insert.append((short_fname, resolved_route, str(agency_val), clean_dr, batch_ts))
                    st.session_state.unmapped_current_batch.append({
                        "File Name": short_fname, "Route": resolved_route, "Agency": agency_val, "Fallback DR": clean_dr
                    })

                is_valid_dr = clean_dr.upper().startswith("DR")
                if is_valid_dr:
                    master_routes_to_insert.append((short_fname, resolved_route, str(agency_val), clean_dr, batch_ts))
                    agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                    seq_num = agency_counts_valid[agency_val]
                    ref_code = f"RT-{resolved_route}-{agency_val}-{today_date}" if seq_num == 1 else f"RT-{resolved_route}-{agency_val}-{today_date}-{seq_num}"
                    target_ws, curr_r, order_id_to_write = ws_valid, valid_r_idx, valid_order_no
                else:
                    agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                    seq_num = agency_counts_missing[agency_val]
                    ref_code = f"RT-{resolved_route}-{agency_val}-{today_date}-NEW" if seq_num == 1 else f"RT-{resolved_route}-{agency_val}-{today_date}-NEW-{seq_num}"
                    target_ws, curr_r, order_id_to_write = ws_missing, missing_r_idx, missing_order_no

                item_seq_id = 10
                row_items_added = 0
                for c_idx, fg_val in valid_cols:
                    q_val = df_input.iloc[r, c_idx]
                    if pd.notna(q_val) and str(q_val).strip() != "":
                        q_str = str(q_val).strip()
                        if q_str.startswith("=") or "SUM(" in q_str.upper():
                            continue
                        try:
                            f_qty = float(q_str)
                            if f_qty > 0:
                                current_fg = fg_val if fg_val.startswith("FG") else col_map_dict.get(c_idx, st.session_state.fg_code)
                                if (agency_val, c_idx) in agency_override_dict:
                                    current_fg = agency_override_dict[(agency_val, c_idx)]

                                cur.execute("""
                                    SELECT id FROM pending_orders 
                                    WHERE route_no=? AND agency_no=? AND fg_code=? AND status='Pending' AND bags_qty=?
                                """, (resolved_route, str(agency_val), current_fg, f_qty))
                                if not cur.fetchone():
                                    pending_records_to_insert.append((
                                        short_fname, f"ORD-{agency_val}-{r}", resolved_route, str(agency_val),
                                        clean_dr, current_fg, f_qty, round(f_qty * 0.05, 2), ref_code, "Pending", batch_ts
                                    ))

                                total_in_qty += f_qty
                                total_gen_qty += f_qty
                                file_input_qty += f_qty
                                row_items_added += 1

                                target_ws.cell(row=curr_r, column=2, value=order_id_to_write)
                                target_ws.cell(row=curr_r, column=3, value="OR")
                                target_ws.cell(row=curr_r, column=4, value="SO20")
                                target_ws.cell(row=curr_r, column=5, value=10)
                                target_ws.cell(row=curr_r, column=6, value=20)
                                target_ws.cell(row=curr_r, column=7, value=clean_dr)
                                target_ws.cell(row=curr_r, column=8, value=clean_dr)
                                target_ws.cell(row=curr_r, column=9, value=ref_code)
                                target_ws.cell(row=curr_r, column=10, value=today_date)
                                target_ws.cell(row=curr_r, column=11, value=today_date)
                                target_ws.cell(row=curr_r, column=15, value=item_seq_id)
                                target_ws.cell(row=curr_r, column=16, value=current_fg)
                                target_ws.cell(row=curr_r, column=19, value=f_qty)
                                target_ws.cell(row=curr_r, column=20, value="Bag")
                                target_ws.cell(row=curr_r, column=22, value=2100)
                                target_ws.cell(row=curr_r, column=26, value=resolved_route)
                                target_ws.cell(row=curr_r, column=27, value=agency_val)

                                item_seq_id += 10
                                curr_r += 1
                        except ValueError:
                            pass

                if row_items_added > 0:
                    if is_valid_dr:
                        valid_r_idx = curr_r
                        valid_order_no += 1
                        valid_items_cnt += row_items_added
                        total_valid += 1
                    else:
                        missing_r_idx = curr_r
                        missing_order_no += 1
                        missing_items_cnt += row_items_added
                        total_missing += 1

            if valid_items_cnt > 0:
                buf_v = io.BytesIO()
                wb_valid.save(buf_v)
                out_name_v = f"{resolved_route}_{today_date}_{time_suffix}_Valid.xlsx"
                st.session_state.processed_files.append({
                    "name": f"{short_fname} (Valid DR)", "data": buf_v.getvalue(), "filename": out_name_v, "orders": valid_items_cnt
                })
                cur.execute(
                    "INSERT OR REPLACE INTO output_files_ledger (file_name, file_type, file_data, created_at) VALUES (?, 'Valid DR', ?, ?)",
                    (out_name_v, buf_v.getvalue(), batch_ts)
                )
                traceability_records.append((batch_ts, short_fname, f_bytes, file_input_qty, out_name_v, "Valid DR", 1, batch_ts))

            if missing_items_cnt > 0:
                buf_m = io.BytesIO()
                wb_missing.save(buf_m)
                out_name_m = f"{resolved_route}_{today_date}_{time_suffix}_Missing_DR.xlsx"
                st.session_state.processed_files.append({
                    "name": f"{short_fname} (Missing DR)", "data": buf_m.getvalue(), "filename": out_name_m, "orders": missing_items_cnt
                })
                cur.execute(
                    "INSERT OR REPLACE INTO output_files_ledger (file_name, file_type, file_data, created_at) VALUES (?, 'Missing DR', ?, ?)",
                    (out_name_m, buf_m.getvalue(), batch_ts)
                )
                traceability_records.append((batch_ts, short_fname, f_bytes, file_input_qty, out_name_m, "Missing DR", 1, batch_ts))

            cur.execute(
                """
                INSERT OR REPLACE INTO uploaded_files_archive (file_name, upload_timestamp, total_records, file_size_kb, batch_status)
                VALUES (?, ?, ?, ?, 'Processed')
            """,
                (short_fname, batch_ts, valid_items_cnt + missing_items_cnt, round(len(f_bytes) / 1024, 2))
            )

        if pending_records_to_insert:
            cur.executemany(
                """
                INSERT INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                pending_records_to_insert
            )

        if master_routes_to_insert:
            cur.executemany(
                "INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)",
                master_routes_to_insert
            )

        if unmapped_records_to_insert:
            cur.executemany(
                "INSERT OR IGNORE INTO unmapped_missing_dr_ledger (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)",
                unmapped_records_to_insert
            )

        if traceability_records:
            cur.executemany(
                """
                INSERT INTO input_output_traceability (batch_timestamp, input_file_name, input_file_blob, total_input_qty, generated_output_file, output_type, version_no, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
                traceability_records
            )

        cur.execute(
            "INSERT INTO history_logs (timestamp, files_count, total_qty, status, remarks) VALUES (?, ?, ?, 'Success', ?)",
            (batch_ts, len(uploaded_files), total_in_qty, f"Clean Qty Extracted: {total_in_qty}")
        )

        conn.commit()
        conn.close()

        st.session_state.kpi_data = {
            "input_qty": total_in_qty, "gen_qty": total_gen_qty, "valid_count": total_valid,
            "missing_count": total_missing, "skipped_count": total_skipped
        }

        if len(st.session_state.processed_files) > 0:
            st.success(f"🎉 Batch execution completed! Clean Qty: {total_in_qty:,.0f} Bags extracted. Created {len(st.session_state.processed_files)} output workbooks & saved {len(pending_records_to_insert)} pending orders.")
        else:
            st.warning("⚠️ Koi valid demand row extract nahi ho saki.")

    # Multi-Channel Export Panel
    if st.session_state.processed_files:
        st.markdown("---")
        st.subheader("📥 Multi-Channel Export & Notification Hub")
        kpi = st.session_state.kpi_data

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in st.session_state.processed_files:
                zf.writestr(item["filename"], item["data"])

        c_exp1, c_exp2, c_exp3, c_exp4, c_exp5, c_exp6, c_exp7 = st.columns(7)
        with c_exp1:
            st.download_button("📦 Download ZIP", zip_buf.getvalue(), f"Batch_Orders_{get_ist_date_str()}.zip", "application/zip")
        with c_exp2:
            try:
                pdf_rep = FPDF()
                pdf_rep.add_page()
                pdf_rep.set_font("Arial", "B", 16)
                pdf_rep.cell(190, 10, "Sales Orders Execution Invoice", ln=True, align="C")
                pdf_rep.set_font("Arial", "", 10)
                pdf_rep.cell(190, 6, f"Generated On (IST): {get_ist_timestamp_full()}", ln=True, align="C")
                pdf_rep.ln(6)
                pdf_rep.set_font("Arial", "B", 10)
                pdf_rep.cell(100, 8, "Metric Description", 1)
                pdf_rep.cell(90, 8, "Value", 1, ln=True)
                pdf_rep.set_font("Arial", "", 10)
                pdf_rep.cell(100, 7, "Total Input Quantity", 1)
                pdf_rep.cell(90, 7, f"{kpi['input_qty']:,.0f} Bags", 1, ln=True)
                pdf_rep.cell(100, 7, "Valid Orders Count", 1)
                pdf_rep.cell(90, 7, str(kpi["valid_count"]), 1, ln=True)
                pdf_rep.cell(100, 7, "Fallback (NEW_CUST) Count", 1)
                pdf_rep.cell(90, 7, str(kpi["missing_count"]), 1, ln=True)
                pdf_rep.cell(100, 7, "Skipped Rows", 1)
                pdf_rep.cell(90, 7, str(kpi["skipped_count"]), 1, ln=True)
                st.download_button("📄 PDF Report", bytes(pdf_rep.output()), f"Execution_Invoice_{get_ist_date_str()}.pdf", "application/pdf")
            except Exception as e:
                st.error(f"PDF Error: {e}")
        with c_exp3:
            txt_summary = f"SALES EXECUTION SUMMARY - {get_ist_timestamp_full()}\nTotal Bags: {kpi['input_qty']}\nValid Orders: {kpi['valid_count']}\nFallback Orders: {kpi['missing_count']}"
            st.download_button("📄 Summary TXT", txt_summary.encode("utf-8"), f"Summary_{get_ist_date_str()}.txt", "text/plain")
        with c_exp4:
            json_dump = json.dumps({"timestamp": get_ist_timestamp_full(), "kpi": kpi}, indent=4)
            st.download_button("💾 Backup JSON", json_dump.encode("utf-8"), f"Audit_{get_ist_date_str()}.json", "application/json")
        with c_exp5:
            components.html("""<button onclick="parent.window.print()" style="width:100%; height:38px; background:#2563eb; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer;">🖨️ Print View</button>""", height=45)
        with c_exp6:
            if st.button("📧 Email Reports"):
                if st.session_state.email_user and st.session_state.email_pass and st.session_state.recipient:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = f"🚀 Sales Demand Execution Report - {get_ist_date_str()}"
                        msg["From"] = st.session_state.email_user
                        msg["To"] = st.session_state.recipient
                        msg.set_content(f"Daily Demand Batch processed successfully on {get_ist_timestamp_full()}.\nTotal Bags: {kpi['input_qty']}")
                        for item in st.session_state.processed_files:
                            msg.add_attachment(item["data"], maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=item["filename"])
                        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                            smtp.login(st.session_state.email_user, st.session_state.email_pass)
                            smtp.send_message(msg)
                        st.success("✅ Email dispatched!")
                    except Exception as e:
                        st.error(f"Email Failed: {e}")
                else:
                    st.warning("⚠️ Enter email credentials in settings!")
        with c_exp7:
            wa_text = f"Sales Batch Processed! Total Bags: {kpi['input_qty']} | Valid: {kpi['valid_count']} | Fallbacks: {kpi['missing_count']}"
            st.markdown(f'<a href="https://wa.me/{st.session_state.whatsapp_num}?text={urllib.parse.quote(wa_text)}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp</button></a>', unsafe_allow_html=True)

        st.markdown("##### Individual File Downloads:")
        for idx_f, f_itm in enumerate(st.session_state.processed_files):
            st.download_button(f"📥 Download {f_itm['name']}", f_itm["data"], f_itm["filename"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_indiv_{idx_f}")

# ==============================================================================
# MODULE 2: ROUTE DISPATCH TRIP PLANNER
# ==============================================================================

elif main_menu == "🚚 Route Dispatch Trip Planner":
    st.title("🚚 Route Dispatch Planning & Vehicle Allocation")
    st.markdown("Intelligent vehicle assignment with **SKU-level partial modifications** and **Overload confirmation protocols**.")

    conn = get_db_connection()
    df_pending = pd.read_sql("SELECT * FROM pending_orders WHERE status='Pending'", conn)

    if df_pending.empty:
        st.info("ℹ️ No pending order demand found. Upload demand workbooks in Module 1.")
    else:
        route_summary = df_pending.groupby("route_no").agg({
            "agency_no": "nunique", "bags_qty": "sum", "weight_mt": "sum"
        }).reset_index().rename(columns={"agency_no": "Total Agencies", "bags_qty": "Total Bags", "weight_mt": "Total MT"})

        st.subheader("📦 Pending Demand Clusters Grouped by Route")
        st.dataframe(route_summary, use_container_width=True)

        st.markdown("---")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            sel_route = st.selectbox("1. Select Route for Dispatch Trip", route_summary["route_no"].tolist())
            route_df = df_pending[df_pending["route_no"] == str(sel_route)].copy()

            avail_fleet = pd.read_sql("SELECT * FROM fleet_master WHERE status='Available'", conn)
            avail_bays = pd.read_sql("SELECT * FROM loading_bays WHERE status='Open'", conn)

            fleet_opts = [f"{r['vehicle_no']} | {r['vehicle_type']} (Cap: {r['capacity_bags']} Bags / {r['capacity_mt']} MT)" for _, r in avail_fleet.iterrows()]
            sel_vehicle = st.selectbox("2. Assign Available Vehicle", fleet_opts if fleet_opts else ["No Vehicles Available"])
            sel_bay = st.selectbox("3. Assign Loading Bay", [f"{r['bay_no']} - {r['bay_name']}" for _, r in avail_bays.iterrows()])

        with col_p2:
            st.markdown("##### 📋 4. Filter Agencies in Route:")
            agencies = route_df["agency_no"].unique().tolist()
            selected_agencies = st.multiselect("Select Agencies to allocate:", agencies, default=agencies)

        st.markdown("---")
        st.subheader("📦 5. Agency & SKU Level Quantity Modification (Include/Exclude/Partial Split):")
        filtered_route_df = route_df[route_df["agency_no"].isin(selected_agencies)].copy()

        filtered_route_df["Include in Trip"] = True
        filtered_route_df["Dispatch Bags"] = filtered_route_df["bags_qty"]

        display_cols = ["id", "Include in Trip", "agency_no", "dr_code", "fg_code", "bags_qty", "Dispatch Bags", "order_no"]

        edited_orders = st.data_editor(
            filtered_route_df[display_cols],
            column_config={
                "id": st.column_config.NumberColumn("Item ID", disabled=True),
                "agency_no": st.column_config.TextColumn("Agency", disabled=True),
                "dr_code": st.column_config.TextColumn("DR Code", disabled=True),
                "fg_code": st.column_config.TextColumn("FG Code", disabled=True),
                "bags_qty": st.column_config.NumberColumn("Original Pending Bags", disabled=True),
                "Include in Trip": st.column_config.CheckboxColumn("Include in Vehicle?"),
                "Dispatch Bags": st.column_config.NumberColumn("Dispatch Qty (Bags)", min_value=0.0, step=1.0)
            },
            hide_index=True,
            use_container_width=True,
            key="trip_sku_editor"
        )

        active_trip_items = edited_orders[edited_orders["Include in Trip"] == True].copy()
        trip_bags = active_trip_items["Dispatch Bags"].sum()
        trip_mt = round(trip_bags * 0.05, 2)

        cap_bags = 0
        is_overloaded = False
        if sel_vehicle != "No Vehicles Available":
            v_num = sel_vehicle.split(" | ")[0]
            v_info = avail_fleet[avail_fleet["vehicle_no"] == v_num].iloc[0]
            cap_bags = int(v_info["capacity_bags"])
            util_pct = (trip_bags / cap_bags * 100) if cap_bags > 0 else 0.0

            c_met1, c_met2 = st.columns(2)
            c_met1.metric("Total Dispatch Bags", f"{trip_bags:,.0f} / {cap_bags} Bags", f"{util_pct:.1f}% Capacity Utilization")
            c_met2.metric("Total Tonnage (MT)", f"{trip_mt:,.2f} MT")

            if trip_bags > cap_bags:
                is_overloaded = True
                st.error(f"🚨 **VEHICLE OVERLOAD WARNING:** Truck capacity exceeded by {trip_bags - cap_bags:,.0f} bags!")
                confirm_overload = st.checkbox("⚠️ Check this box to CONFIRM and OVERRIDE vehicle capacity limits for this trip.")
            else:
                confirm_overload = True

        st.markdown("---")
        submit_disabled = (sel_vehicle == "No Vehicles Available") or active_trip_items.empty or (is_overloaded and not confirm_overload)

        if st.button("🚀 Confirm Trip & Generate Loading Slip", type="primary", disabled=submit_disabled):
            cur = conn.cursor()
            now_ist = get_ist_now()
            trip_id = f"TRIP-{sel_route}-{now_ist.strftime('%Y%m%d%H%M%S')}"
            v_num = sel_vehicle.split(" | ")[0]
            v_info = avail_fleet[avail_fleet["vehicle_no"] == v_num].iloc[0]
            bay_code = sel_bay.split(" - ")[0]

            cur.execute("""
                INSERT INTO trip_loading_slips (trip_id, trip_date, route_no, vehicle_no, transporter_name, driver_name, driver_phone, loading_bay, total_bags, total_weight_mt, capacity_utilization_pct, status, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Planned', ?)
            """, (trip_id, now_ist.strftime("%Y-%m-%d"), str(sel_route), v_num, v_info["transporter_name"], v_info["driver_name"], v_info["driver_phone"], bay_code, trip_bags, trip_mt, round((trip_bags/v_info["capacity_bags"]*100), 2), now_ist.strftime("%Y-%m-%d %H:%M:%S")))

            seq = 1
            for _, r_val in edited_orders.iterrows():
                item_id = r_val["id"]
                orig_qty = float(r_val["bags_qty"])
                inc = r_val["Include in Trip"]
                disp_qty = float(r_val["Dispatch Bags"]) if inc else 0.0

                item_row = df_pending[df_pending["id"] == item_id].iloc[0]

                if inc and disp_qty > 0:
                    cur.execute("""
                        INSERT INTO trip_order_items (trip_id, order_no, agency_no, route_no, dr_code, fg_code, allocated_bags, allocated_weight_mt, delivery_seq, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Assigned')
                    """, (trip_id, item_row["order_no"], item_row["agency_no"], item_row["route_no"], item_row["dr_code"], item_row["fg_code"], disp_qty, round(disp_qty*0.05, 2), seq))
                    seq += 1

                rem_qty = orig_qty - disp_qty
                if rem_qty > 0:
                    cur.execute("UPDATE pending_orders SET bags_qty=?, weight_mt=?, status='Pending' WHERE id=?", (rem_qty, round(rem_qty*0.05, 2), item_id))
                    cur.execute("""
                        INSERT INTO partial_dispatch_ledger (trip_id, source_file, order_no, route_no, agency_no, dr_code, fg_code, original_bags, dispatched_bags, remaining_bags, status, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Partial Pending', ?)
                    """, (trip_id, item_row["source_file"], item_row["order_no"], item_row["route_no"], item_row["agency_no"], item_row["dr_code"], item_row["fg_code"], orig_qty, disp_qty, rem_qty, now_ist.strftime("%Y-%m-%d %H:%M:%S")))
                elif rem_qty == 0 and inc:
                    cur.execute("UPDATE pending_orders SET status='Assigned' WHERE id=?", (item_id,))

            cur.execute("UPDATE fleet_master SET status='Assigned to Trip' WHERE vehicle_no=?", (v_num,))
            conn.commit()
            st.success(f"🎉 Trip '{trip_id}' Created Successfully!")
            st.rerun()
    conn.close()

# ==============================================================================
# MODULE 3: LIVE INVENTORY STOCK & ERP DEMAND MATCHER
# ==============================================================================

elif main_menu == "📦 Live Inventory Stock & ERP Demand Matcher":
    st.title("📦 Live Plant Stock & Inventory Ledger")
    conn = get_db_connection()
    df_stk = pd.read_sql("SELECT * FROM plant_inventory_stock", conn)
    st.dataframe(df_stk, use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 4: LOADING SLIPS & ACTIVE TRIPS
# ==============================================================================

elif main_menu == "📋 Loading Slips & Active Trips":
    st.title("📋 Trip Loading Slips & Active Trips")
    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips ORDER BY created_at DESC", conn)
    st.dataframe(df_trips, use_container_width=True)
    if not df_trips.empty:
        sel_trip = st.selectbox("Select Trip ID to inspect:", df_trips["trip_id"].tolist())
        trip_row = df_trips[df_trips["trip_id"] == sel_trip].iloc[0]
        items_df = pd.read_sql("SELECT * FROM trip_order_items WHERE trip_id=?", conn, params=(sel_trip,))
        st.dataframe(items_df, use_container_width=True)
        pdf_bytes = build_pdf_loading_slip(trip_row.to_dict(), items_df)
        st.download_button("📄 Download PDF Loading Slip", pdf_bytes, f"Loading_Slip_{sel_trip}.pdf", "application/pdf")
    conn.close()

# ==============================================================================
# MODULE 5: DAILY DISPATCH SALE REGISTER
# ==============================================================================

elif main_menu == "📖 Daily Dispatch Sale Register":
    st.title("📖 Daily Dispatch Sale Register")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM daily_dispatch_register", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 6: PARTIAL / SPLIT DISPATCH DATABASE
# ==============================================================================

elif main_menu == "🧩 Partial / Split Dispatch Database":
    st.title("🧩 Partial / Split Dispatch Database")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM partial_dispatch_ledger", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 7: PENDING ORDERS LEDGER
# ==============================================================================

elif main_menu == "⏳ Pending Orders Ledger":
    st.title("⏳ Pending Orders Ledger")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM pending_orders", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 8: FILE UPLOAD ARCHIVE
# ==============================================================================

elif main_menu == "🗄️ File Upload Archive":
    st.title("🗄️ Uploaded Input File Archive")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM uploaded_files_archive", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 9: MASTER DB & UNMAPPED LEDGER
# ==============================================================================

elif main_menu == "📋 Master DB & Unmapped Ledger":
    st.title("📋 Master DB & Unmapped Ledger")
    conn = get_db_connection()
    t1, t2 = st.tabs(["Master DB", "Unmapped Ledger"])
    with t1: st.dataframe(pd.read_sql("SELECT * FROM unique_routes_master", conn), use_container_width=True)
    with t2: st.dataframe(pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 10: FLEET & LOADING BAY MASTER
# ==============================================================================

elif main_menu == "🚛 Fleet & Loading Bay Master":
    st.title("🚛 Fleet & Loading Bay Master")
    conn = get_db_connection()
    f1, f2 = st.tabs(["Fleet Master", "Loading Bays"])
    with f1: st.dataframe(pd.read_sql("SELECT * FROM fleet_master", conn), use_container_width=True)
    with f2: st.dataframe(pd.read_sql("SELECT * FROM loading_bays", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 11: TRACEABILITY & AUDIT LEDGERS
# ==============================================================================

elif main_menu == "🔍 Traceability & Audit Ledgers":
    st.title("🔍 Traceability & Audit Ledgers")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM input_output_traceability", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 12: EXECUTIVE KPI & VISUAL ANALYTICS
# ==============================================================================

elif main_menu == "📊 Executive KPI & Visual Analytics":
    st.title("📊 Supply Chain & Dispatch KPI Analytics Dashboard")
    conn = get_db_connection()
    df_t = pd.read_sql("SELECT * FROM trip_loading_slips", conn)
    st.metric("Total Trips Planned", len(df_t))
    conn.close()

# ==============================================================================
# MODULE 13: UNIVERSAL DATE & MULTI-FIELD FILTER CENTER
# ==============================================================================

elif main_menu == "🎯 Universal Date & Multi-Field Filter Center":
    st.title("🎯 Universal Date & Multi-Field Filter Center")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    tbls = [r[0] for r in cur.fetchall()]
    sel_t = st.selectbox("Select Table to Inspect:", tbls)
    if sel_t:
        st.dataframe(pd.read_sql(f"SELECT * FROM {sel_t}", conn), use_container_width=True)
    conn.close()

# ==============================================================================
# MODULE 14: SMART MULTI-TRUCK LOAD OPTIMIZER PRO
# ==============================================================================

elif main_menu == "⚡ Smart Multi-Truck Load Optimizer Pro":
    st.title("⚡ Smart Multi-Truck Load Optimizer Pro")
    st.markdown("Automated Bin-Packing Algorithm for Route Truck Load Distribution.")

# ==============================================================================
# MODULE 15: IN-APP DATABASE BUILDER & DYNAMIC LINKER
# ==============================================================================

elif main_menu == "🗄️ In-App Database Builder & Dynamic Linker":
    st.title("🗄️ In-App Dynamic Database Builder & Universal CRUD")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    st.write([r[0] for r in cur.fetchall()])
    conn.close()
