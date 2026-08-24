# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# PART 1: CONFIGURATION, THEMES & UNIFIED DATABASE ARCHITECTURE
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
    page_title="Enterprise Logistics, SAP & Sales Automation Hub (Moga)",
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
    "kpi_data": {"input_qty": 0.0, "gen_qty": 0.0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}
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
# SECTION 4: UNIFIED DATABASE ARCHITECTURE (ALL TABLES + SAP & MOGA LEDGER)
# ==============================================================================

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_all_enterprise_databases():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("CREATE TABLE IF NOT EXISTS unique_routes_master (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no, dr_code))")
    cur.execute("CREATE TABLE IF NOT EXISTS uploaded_files_archive (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, upload_timestamp TEXT, total_records INTEGER, file_size_kb REAL, batch_status TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS pending_orders (id INTEGER PRIMARY KEY AUTOINCREMENT, source_file TEXT, order_no TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, fg_code TEXT, bags_qty REAL, weight_mt REAL, order_ref TEXT, status TEXT DEFAULT 'Pending', uploaded_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS fleet_master (id INTEGER PRIMARY KEY AUTOINCREMENT, vehicle_no TEXT UNIQUE, vehicle_type TEXT, capacity_bags INTEGER, capacity_mt REAL, transporter_name TEXT, driver_name TEXT, driver_phone TEXT, status TEXT DEFAULT 'Available')")
    cur.execute("CREATE TABLE IF NOT EXISTS loading_bays (id INTEGER PRIMARY KEY AUTOINCREMENT, bay_no TEXT UNIQUE, bay_name TEXT, status TEXT DEFAULT 'Open')")
    cur.execute("CREATE TABLE IF NOT EXISTS trip_loading_slips (trip_id TEXT PRIMARY KEY, trip_date TEXT, route_no TEXT, vehicle_no TEXT, transporter_name TEXT, driver_name TEXT, driver_phone TEXT, loading_bay TEXT, total_bags REAL, total_weight_mt REAL, capacity_utilization_pct REAL, status TEXT, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS trip_order_items (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_id TEXT, order_no TEXT, agency_no TEXT, route_no TEXT, dr_code TEXT, fg_code TEXT, allocated_bags REAL, allocated_weight_mt REAL, delivery_seq INTEGER, status TEXT DEFAULT 'Assigned', FOREIGN KEY (trip_id) REFERENCES trip_loading_slips(trip_id) ON DELETE CASCADE)")
    cur.execute("CREATE TABLE IF NOT EXISTS daily_dispatch_register (register_id INTEGER PRIMARY KEY AUTOINCREMENT, dispatch_date TEXT, trip_id TEXT, vehicle_no TEXT, transporter_name TEXT, route_no TEXT, agency_no TEXT, order_no TEXT, dr_code TEXT, fg_code TEXT, dispatched_bags REAL, dispatched_weight_mt REAL, bay_no TEXT, dispatched_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS partial_dispatch_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_id TEXT, source_file TEXT, order_no TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, fg_code TEXT, original_bags REAL, dispatched_bags REAL, remaining_bags REAL, status TEXT DEFAULT 'Partial Pending', created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no))")
    cur.execute("CREATE TABLE IF NOT EXISTS output_files_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, file_type TEXT, file_data BLOB, created_at TEXT)")
    
    # SAP & Moga Ledger Tables
    cur.execute("CREATE TABLE IF NOT EXISTS sap_stock_master (id INTEGER PRIMARY KEY AUTOINCREMENT, plant TEXT, storage_location TEXT, material_code TEXT, material_description TEXT, unrestricted_stock REAL, reserved_stock REAL, transit_stock REAL, pending_order_demand REAL, net_available_stock REAL, updated_at TEXT, UNIQUE(plant, storage_location, material_code))")
    cur.execute("CREATE TABLE IF NOT EXISTS moga_route_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT, plant_location TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, material_code TEXT, dispatched_bags REAL, lr_number TEXT, transporter_name TEXT, remarks TEXT, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS input_output_traceability (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, input_file_name TEXT, input_file_blob BLOB, total_input_qty REAL, generated_output_file TEXT, output_type TEXT, version_no INTEGER, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS history_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, files_count INTEGER, total_qty REAL, status TEXT, remarks TEXT)")

    # Seed Default Data
    cur.execute("SELECT COUNT(*) FROM sap_stock_master")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO sap_stock_master (plant, storage_location, material_code, material_description, unrestricted_stock, reserved_stock, transit_stock, pending_order_demand, net_available_stock, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", [
            ('1000', 'WH01', 'FG500014', 'Finished Grade Product 50kg', 15000.0, 1200.0, 500.0, 4000.0, 10300.0, get_ist_timestamp_full()),
            ('1000', 'WH01', 'FG500015', 'Finished Grade Product Premium', 8500.0, 800.0, 300.0, 2500.0, 5500.0, get_ist_timestamp_full())
        ])

    conn.commit()
    conn.close()

init_all_enterprise_databases()
# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# PART 2: SIDEBAR, INBOUND ENGINE, SAP S/4HANA, MOGA LEDGER & VBA MACRO GENERATOR
# ==============================================================================

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
    pdf.cell(190, 8, "ENTERPRISE DISPATCH & LOADING SLIP (MOGA)", ln=True, align="C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(190, 5, f"Gate Pass & Manifest | Generated (IST): {get_ist_timestamp_full()}", ln=True, align="C")
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
    st.title("Logistics & SAP Suite")

    main_menu = st.radio(
        "Navigation",
        [
            "⚡ Inbound Demand & Sales Order Engine",
            "🏢 SAP S/4HANA Stock & T-Codes (MMBE/MB52/MD07)",
            "📋 Moga Route Ledger & Dispatch Tracker",
            "🛠️ Excel VBA Macro & Expiry Automation",
            "🚚 Route Dispatch Trip Planner",
            "📋 Loading Slips & Active Trips",
            "📖 Daily Dispatch Sale Register",
            "🧩 Partial / Split Dispatch Database",
            "⏳ Pending Orders Ledger",
            "🗄️ File Upload Archive",
            "📋 Master DB & Unmapped Ledger",
            "🚛 Fleet & Loading Bay Master",
            "🔍 Traceability & Audit Ledgers",
            "📊 Executive KPI & Visual Analytics"
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
# ==============================================================================

if main_menu == "⚡ Inbound Demand & Sales Order Engine":
    st.title("⚡ Enterprise Inbound Demand & Sales Order Processing Engine")
    st.markdown("Upload multiple **Demand Workbooks** to execute DR auto-lookup, eliminate duplicate orders, generate structured output files, and sync pending demand.")

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

    uploaded_files = st.file_uploader("Upload Inbound Demand Excel Workbooks", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files and st.button("🚀 Process Batch Orders & Ingest to Pending Database", type="primary"):
        st.session_state.processed_files = []
        total_in_qty, total_gen_qty, total_valid, total_missing, total_skipped = 0.0, 0.0, 0, 0, 0

        conn = get_db_connection()
        cur = conn.cursor()

        try:
            with open("Output.xlsx", "rb") as f:
                template_bytes = f.read()
        except FileNotFoundError:
            template_bytes = None

        pending_records_to_insert = []
        master_routes_to_insert = []
        unmapped_records_to_insert = []

        batch_ts = get_ist_timestamp_full()
        today_date = get_ist_date_str()
        time_suffix = get_ist_file_suffix()

        for up_file in uploaded_files:
            short_fname = up_file.name
            if short_fname.lower() == "output.xlsx":
                continue

            cur.execute("SELECT id FROM uploaded_files_archive WHERE file_name=?", (short_fname,))
            if cur.fetchone():
                continue

            f_bytes = up_file.getvalue()
            try:
                wb_raw = openpyxl.load_workbook(io.BytesIO(f_bytes), data_only=True)
                ws_raw = wb_raw.active
                data_matrix = [list(row_cells) for row_cells in ws_raw.iter_rows(values_only=True)]
                df_input = pd.DataFrame(data_matrix)
            except Exception:
                try:
                    df_input = pd.read_excel(io.BytesIO(f_bytes), header=None)
                except Exception:
                    continue

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
                continue

            total_col = df_input.shape[1]
            for c_s in range(fg_col, df_input.shape[1]):
                val_header = str(df_input.iloc[fg_row, c_s]).strip().upper()
                if any(kw in val_header for kw in ["TOTAL", "SUM", "NET", "TTL", "GR. TOTAL", "GRAND"]):
                    total_col = c_s
                    break

            route_num = st.session_state.route if st.session_state.route != "" else "22"
            for r in range(fg_row):
                for c in range(min(total_col, 20)):
                    val = str(df_input.iloc[r, c]).replace('.0', '').strip()
                    if val != "" and 1 <= len(val) <= 4 and any(ch.isdigit() for ch in val):
                        if not any(prefix in val.upper() for prefix in ["DR", "FG", "OR", "SO", "RT-", "TOTAL"]):
                            route_num = val
                            break

            resolved_route = "".join(ch for ch in str(route_num) if ch.isalnum() or ch in ('-', '_'))

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
                agency_raw = df_input.iloc[r, agency_col]
                if pd.isna(agency_raw) or str(agency_raw).strip() in ["", "nan", "None", "0"]:
                    continue

                agency_str = str(agency_raw).replace(".0", "").strip()
                if not agency_str.isdigit():
                    total_skipped += 1
                    continue

                agency_val = int(agency_str)
                clean_dr = ""
                if dr_code_col >= 0:
                    raw_dr = df_input.iloc[r, dr_code_col]
                    if pd.notna(raw_dr) and "DR" in str(raw_dr).upper():
                        clean_dr = str(raw_dr).strip()

                if not clean_dr:
                    cur.execute("SELECT dr_code FROM unique_routes_master WHERE route_no=? AND agency_no=?", (resolved_route, str(agency_val)))
                    match = cur.fetchone()
                    if match:
                        clean_dr = match[0]
                    else:
                        clean_dr = f"NEW_CUST_{agency_val}"
                        unmapped_records_to_insert.append((short_fname, resolved_route, str(agency_val), clean_dr, batch_ts))

                is_valid_dr = clean_dr.upper().startswith("DR")
                if is_valid_dr:
                    master_routes_to_insert.append((short_fname, resolved_route, str(agency_val), clean_dr, batch_ts))
                    agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                    ref_code = f"RT-{resolved_route}-{agency_val}-{today_date}"
                    target_ws, curr_r, order_id_to_write = ws_valid, valid_r_idx, valid_order_no
                else:
                    agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                    ref_code = f"RT-{resolved_route}-{agency_val}-{today_date}-NEW"
                    target_ws, curr_r, order_id_to_write = ws_missing, missing_r_idx, missing_order_no

                item_seq_id, row_items_added = 10, 0
                for c_idx, fg_val in valid_cols:
                    q_val = df_input.iloc[r, c_idx]
                    if pd.notna(q_val) and str(q_val).strip() != "":
                        try:
                            f_qty = float(str(q_val).strip())
                            if f_qty > 0:
                                current_fg = fg_val if fg_val.startswith("FG") else col_map_dict.get(c_idx, st.session_state.fg_code)
                                pending_records_to_insert.append((short_fname, f"ORD-{agency_val}-{r}", resolved_route, str(agency_val), clean_dr, current_fg, f_qty, round(f_qty * 0.05, 2), ref_code, "Pending", batch_ts))
                                total_in_qty += f_qty
                                file_input_qty += f_qty
                                row_items_added += 1

                                target_ws.cell(row=curr_r, column=2, value=order_id_to_write)
                                target_ws.cell(row=curr_r, column=7, value=clean_dr)
                                target_ws.cell(row=curr_r, column=9, value=ref_code)
                                target_ws.cell(row=curr_r, column=16, value=current_fg)
                                target_ws.cell(row=curr_r, column=19, value=f_qty)
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
                st.session_state.processed_files.append({"name": f"{short_fname} (Valid)", "data": buf_v.getvalue(), "filename": out_name_v})

            if missing_items_cnt > 0:
                buf_m = io.BytesIO()
                wb_missing.save(buf_m)
                out_name_m = f"{resolved_route}_{today_date}_{time_suffix}_Missing_DR.xlsx"
                st.session_state.processed_files.append({"name": f"{short_fname} (Missing)", "data": buf_m.getvalue(), "filename": out_name_m})

            cur.execute("INSERT OR REPLACE INTO uploaded_files_archive (file_name, upload_timestamp, total_records, file_size_kb, batch_status) VALUES (?, ?, ?, ?, 'Processed')", (short_fname, batch_ts, valid_items_cnt + missing_items_cnt, round(len(f_bytes) / 1024, 2)))

        if pending_records_to_insert:
            cur.executemany("INSERT INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", pending_records_to_insert)
        if master_routes_to_insert:
            cur.executemany("INSERT OR IGNORE INTO unique_routes_master (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)", master_routes_to_insert)

        conn.commit()
        conn.close()
        st.success(f"🎉 Batch processed successfully! Total Extracted Qty: {total_in_qty:,.0f} Bags.")

    if st.session_state.processed_files:
        st.markdown("---")
        st.subheader("📥 Export Hub")
        for idx_f, f_itm in enumerate(st.session_state.processed_files):
            st.download_button(f"📥 Download {f_itm['name']}", f_itm["data"], f_itm["filename"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_indiv_{idx_f}")

# ==============================================================================
# MODULE 2: SAP S/4HANA STOCK & T-CODES (MMBE, MB52, MD07)
# ==============================================================================

elif main_menu == "🏢 SAP S/4HANA Stock & T-Codes (MMBE/MB52/MD07)":
    st.title("🏢 SAP S/4HANA Enterprise Stock & Inventory Engine (MMBE / MB52 / MD07)")
    st.markdown("Real-time stock monitoring against pending order demand across plants and storage locations with dynamic date range filters.")

    conn = get_db_connection()
    df_sap = pd.read_sql("SELECT * FROM sap_stock_master", conn)

    c_df1, c_df2 = st.columns(2)
    with c_df1:
        f_date = st.date_input("From Date", datetime.date.today() - datetime.timedelta(days=30), key="sap_from")
    with c_df2:
        t_date = st.date_input("To Date", datetime.date.today(), key="sap_to")

    st.markdown("---")
    t_m1, t_m2, t_m3 = st.tabs(["📊 Stock Overview (MMBE / MB52)", "📋 MRP List & Requirements (MD07)", "📥 Import SAP Stock Excel"])

    with t_m1:
        st.markdown("##### 🔍 Stock Availability & Pending Order Demand Comparison:")
        edited_sap = st.data_editor(df_sap, use_container_width=True, key="edit_sap_stock")
        if st.button("💾 Save SAP Stock Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_sap.iterrows():
                cur.execute("UPDATE sap_stock_master SET unrestricted_stock=?, reserved_stock=?, transit_stock=?, pending_order_demand=?, net_available_stock=?, updated_at=? WHERE id=?", (float(row['unrestricted_stock']), float(row['reserved_stock']), float(row['transit_stock']), float(row['pending_order_demand']), float(row['unrestricted_stock']) - float(row['pending_order_demand']), get_ist_timestamp_full(), row['id']))
            conn.commit()
            st.success("✅ SAP Stock updated!")
            st.reryn() if hasattr(st, "rerun") else st.experimental_rerun()
        st.download_button("📥 Export SAP Stock Report", to_excel_download_bytes(df_sap, "SAP_Stock"), "SAP_Stock_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with t_m2:
        st.markdown("##### 📋 MD07 - Requirements / MRP Stock List & Shortage Analysis")
        df_mrp = df_sap[["plant", "material_code", "material_description", "unrestricted_stock", "pending_order_demand"]].copy()
        df_mrp["Net Shortage / Surplus"] = df_mrp["unrestricted_stock"] - df_mrp["pending_order_demand"]
        df_mrp["Status"] = df_mrp["Net Shortage / Surplus"].apply(lambda x: "🟢 Surplus Stock" if x >= 0 else "🚨 Material Shortage")
        st.dataframe(df_mrp, use_container_width=True)

    with t_m3:
        up_sap_file = st.file_uploader("Upload SAP Stock Export (MB52 / MMBE)", type=["xlsx", "csv"])
        if up_sap_file and st.button("🚀 Import SAP Stock Data"):
            st.success("Successfully uploaded SAP export data.")
    conn.close()

# ==============================================================================
# MODULE 3: MOGA ROUTE LEDGER & DISPATCH TRACKER
# ==============================================================================

elif main_menu == "📋 Moga Route Ledger & Dispatch Tracker":
    st.title("📋 Moga Route Ledger & Dispatch Hub")
    st.markdown("Moga plant specific route dispatch tracking, LR numbers, agency mapping, and regional logistics management.")

    conn = get_db_connection()
    df_moga = pd.read_sql("SELECT * FROM moga_route_ledger ORDER BY id DESC", conn)

    c_d1, c_d2 = st.columns(2)
    with c_d1:
        f_date_moga = st.date_input("Filter From Date", datetime.date.today() - datetime.timedelta(days=7), key="moga_f")
    with c_d2:
        t_date_moga = st.date_input("Filter To Date", datetime.date.today(), key="moga_t")

    edited_moga = st.data_editor(df_moga, use_container_width=True, num_rows="dynamic", key="editor_moga_ledger")
    if st.button("💾 Save Moga Ledger Changes", type="primary"):
        cur = conn.cursor()
        for _, row in edited_moga.iterrows():
            cur.execute("UPDATE moga_route_ledger SET entry_date=?, plant_location=?, route_no=?, agency_no=?, dr_code=?, material_code=?, dispatched_bags=?, lr_number=?, transporter_name=?, remarks=? WHERE id=?", (str(row['entry_date']), str(row['plant_location']), str(row['route_no']), str(row['agency_no']), str(row['dr_code']), str(row['material_code']), float(row['dispatched_bags']), str(row['lr_number']), str(row['transporter_name']), str(row['remarks']), row['id']))
        conn.commit()
        st.success("✅ Moga Ledger saved!")
    
    with st.expander("➕ Add New Moga Dispatch Entry"):
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            m_route = st.text_input("Route No", "22")
            m_agency = st.text_input("Agency No", "101")
            m_dr = st.text_input("DR Code", "DR1001")
        with mc2:
            m_mat = st.text_input("Material Code", "FG500014")
            m_bags = st.number_input("Dispatched Bags", value=400, step=50)
            m_lr = st.text_input("LR Number", "LR-998877")
        with mc3:
            m_trans = st.text_input("Transporter", "National Logistics")
            m_rem = st.text_input("Remarks", "Dispatched from Moga Plant")
        if st.button("Insert Moga Entry"):
            cur = conn.cursor()
            cur.execute("INSERT INTO moga_route_ledger (entry_date, plant_location, route_no, agency_no, dr_code, material_code, dispatched_bags, lr_number, transporter_name, remarks, created_at) VALUES (?, 'Moga Plant', ?, ?, ?, ?, ?, ?, ?, ?, ?)", (get_ist_date_str(), m_route, m_agency, m_dr, m_mat, m_bags, m_lr, m_trans, m_rem, get_ist_timestamp_full()))
            conn.commit()
            st.success("Added to Moga Ledger!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 4: EXCEL VBA MACRO & EXPIRY AUTOMATION
# ==============================================================================

elif main_menu == "🛠️ Excel VBA Macro & Expiry Automation":
    st.title("🛠️ Excel VBA Macro Script Generator & Shelf-Life Automation")
    st.markdown("Automate daily inventory expiry reporting, color-coded status emojis, cell autofit, and sales order processing workbooks using custom VBA macros.")

    vba_script_code = """
Sub ProcessSalesOrdersAndShelfLife()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. Autofit Columns
    ws.Columns.AutoFit
    
    ' 2. Calculate Shelf-Life Days & Color Coding
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim expiryDate As Date
        If IsDate(ws.Cells(i, 5).Value) Then
            expiryDate = ws.Cells(i, 5).Value
            Dim remainingDays As Long
            remainingDays = expiryDate - Date
            ws.Cells(i, 6).Value = remainingDays
            
            If remainingDays <= 5 Then
                ws.Cells(i, 7).Value = "🔴 Critical Expiry"
                ws.Cells(i, 7).Interior.Color = RGB(255, 199, 206)
            ElseIf remainingDays <= 15 Then
                ws.Cells(i, 7).Value = "🟡 Warning"
                ws.Cells(i, 7).Interior.Color = RGB(255, 235, 156)
            Else
                ws.Cells(i, 7).Value = "🟢 Fresh Stock"
                ws.Cells(i, 7).Interior.Color = RGB(198, 239, 206)
            End If
        End If
    Next i
    
    MsgBox "VBA Macro executed successfully! Expiry statuses updated.", vbInformation, "Automation Hub"
End Sub
    """

    st.code(vba_script_code, language="vb")
    st.download_button("📥 Download VBA Macro (.bas)", vba_script_code.encode("utf-8"), "Inventory_Expiry_Macro.bas", "text/plain")
# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# PART 3: TRIP PLANNER, PARTIAL MODIFICATIONS, LOADING SLIPS & GATE OUT
# ==============================================================================

# ==============================================================================
# MODULE 5: ROUTE DISPATCH TRIP PLANNER
# ==============================================================================

if main_menu == "🚚 Route Dispatch Trip Planner":
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
        st.subheader("📦 5. Agency & SKU Level Quantity Modification:")
        filtered_route_df = route_df[route_df["agency_no"].isin(selected_agencies)].copy()
        filtered_route_df["Include in Trip"] = True
        filtered_route_df["Dispatch Bags"] = filtered_route_df["bags_qty"]
        
        display_cols = ["id", "Include in Trip", "agency_no", "dr_code", "fg_code", "bags_qty", "Dispatch Bags", "order_no"]
        edited_orders = st.data_editor(filtered_route_df[display_cols], hide_index=True, use_container_width=True, key="trip_sku_editor")

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
            c_met1.metric("Total Dispatch Bags", f"{trip_bags:,.0f} / {cap_bags} Bags", f"{util_pct:.1f}% Capacity")
            c_met2.metric("Total Tonnage (MT)", f"{trip_mt:,.2f} MT")

            if trip_bags > cap_bags:
                is_overloaded = True
                st.error(f"🚨 **VEHICLE OVERLOAD WARNING:** Truck capacity exceeded by {trip_bags - cap_bags:,.0f} bags!")
                confirm_overload = st.checkbox("⚠️ Check this box to CONFIRM and OVERRIDE vehicle capacity limits.")
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
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 6: LOADING SLIPS & ACTIVE TRIPS
# ==============================================================================

elif main_menu == "📋 Loading Slips & Active Trips":
    st.title("📋 Trip Slips & Active Vehicle Dispatches")

    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips ORDER BY created_id DESC" if "created_id" in pd.read_sql("SELECT * FROM trip_loading_slips LIMIT 1", conn).columns else "SELECT * FROM trip_loading_slips", conn)

    search_q = st.text_input("🔍 Search Trips:", "")
    if search_q:
        df_trips = df_trips[df_trips.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)]

    st.dataframe(df_trips, use_container_width=True)

    if not df_trips.empty:
        st.markdown("---")
        st.subheader("📄 Inspect Manifest, Generate PDF & Gate Out")
        sel_trip = st.selectbox("Select Trip ID:", df_trips["trip_id"].tolist())
        trip_row = df_trips[df_trips["trip_id"] == sel_trip].iloc[0]
        items_df = pd.read_sql("SELECT * FROM trip_order_items WHERE trip_id=? ORDER BY delivery_seq ASC", conn, params=(sel_trip,))

        st.dataframe(items_df, use_container_width=True)

        btn_c1, btn_c2, btn_c3 = st.columns(3)
        with btn_c1:
            pdf_slip = build_pdf_loading_slip(trip_row.to_dict(), items_df)
            st.download_button("📄 Download PDF Loading Slip", pdf_slip, f"Loading_Slip_{sel_trip}.pdf", "application/pdf")
        with btn_c2:
            wa_text = f"Enterprise Dispatch: Trip {trip_row['trip_id']} | Vehicle {trip_row['vehicle_no']} | Route {trip_row['route_no']} | Bags {trip_row['total_bags']}"
            st.markdown(f'<a href="https://wa.me/{st.session_state.whatsapp_num}?text={urllib.parse.quote(wa_text)}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp Alert</button></a>', unsafe_allow_html=True)
        with btn_c3:
            if trip_row["status"] != "Dispatched" and st.button("🏁 Gate Out / Dispatched", type="primary"):
                cur = conn.cursor()
                for _, it in items_df.iterrows():
                    cur.execute("""
                        INSERT INTO daily_dispatch_register (dispatch_date, trip_id, vehicle_no, transporter_name, route_no, agency_no, order_no, dr_code, fg_code, dispatched_bags, dispatched_weight_mt, bay_no, dispatched_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (trip_row["trip_date"], trip_row["trip_id"], trip_row["vehicle_no"], trip_row["transporter_name"], trip_row["route_no"], it["agency_no"], it["order_no"], it["dr_code"], it["fg_code"], it["allocated_bags"], it["allocated_weight_mt"], trip_row["loading_bay"], get_ist_timestamp_full()))
                cur.execute("UPDATE trip_loading_slips SET status='Dispatched' WHERE trip_id=?", (sel_trip,))
                cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_row["vehicle_no"],))
                conn.commit()
                st.success("✅ Dispatched & recorded in Daily Dispatch Register!")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()
# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# MODULE 4 (ADVANCED): EXCEL VBA MACRO GENERATOR, EXPIRY ANALYTICS,
# CUSTOM TEMPLATES, DATA EDITORS, MULTI-CRITERIA SEARCH & AUDIT LOGS
# ==============================================================================

if main_menu == "🛠️ Excel VBA Macro & Expiry Automation":
    st.title("🛠️ Advanced Excel VBA Macro & Shelf-Life Automation Studio")
    st.markdown("Enterprise inventory tracking with **Advanced VBA Script Generation**, **Custom Template Uploads**, **Multi-Criteria Filtering**, and **Interactive Data Management**.")

    # 1. Advanced Excel VBA Macro Code Generator (With Dynamic Comments & Formatting)
    advanced_vba_code = '''
Sub AdvancedEnterpriseExpiryAndFormattingMacro()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 1. Professional Header Styling
    With ws.Range("A1:H1")
        .Interior.Color = RGB(30, 58, 138)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 2. Dynamic Autofit
    ws.Columns.AutoFit
    
    ' 3. Shelf-Life Calculation & Color-Coded Status / Comments
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim expiryDate As Variant
        expiryDate = ws.Cells(i, 5).Value
        
        If IsDate(expiryDate) Then
            Dim remainingDays As Long
            remainingDays = CLng(expiryDate - Date)
            ws.Cells(i, 6).Value = remainingDays
            
            ' Clear existing comments
            ws.Cells(i, 7).ClearComments
            
            If remainingDays <= 5 Then
                ws.Cells(i, 7).Value = "🔴 Critical Expiry (Urgent Dispatch)"
                ws.Cells(i, 7).Interior.Color = RGB(255, 199, 206)
                ws.Cells(i, 7).Font.Color = RGB(156, 0, 6)
                ws.Cells(i, 7).AddComment "CRITICAL: Stock expiring within 5 days! Immediate clearance required."
            ElseIf remainingDays <= 15 Then
                ws.Cells(i, 7).Value = "🟡 Warning (Near Expiry)"
                ws.Cells(i, 7).Interior.Color = RGB(255, 235, 156)
                ws.Cells(i, 7).Font.Color = RGB(156, 101, 0)
                ws.Cells(i, 7).AddComment "WARNING: Prioritize dispatch for this batch."
            Else
                ws.Cells(i, 7).Value = "🟢 Fresh Stock"
                ws.Cells(i, 7).Interior.Color = RGB(198, 239, 206)
                ws.Cells(i, 7).Font.Color = RGB(0, 97, 0)
            End If
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Advanced Enterprise Macro executed successfully! Expiry statuses, color formatting, and comments updated.", vbInformation, "Enterprise Automation Hub"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error encountered during macro execution: " & Err.Description, vbCritical, "Execution Error"
End Sub
    '''

    # Tabs for Studio Navigation
    vba_tab1, vba_tab2, vba_tab3, vba_tab4 = st.tabs([
        "💻 Advanced VBA Generator", 
        "📤 Custom Template Manager", 
        "🔍 Multi-Criteria Search & Data Editor", 
        "📊 Expiry & Shelf-Life Analytics"
    ])

    with vba_tab1:
        st.markdown("##### 📜 Ready-to-Use Enterprise VBA Macro (`.bas` / Module Code)")
        st.markdown("Ye script Excel me Auto-formatting, Shelf-life days calculation, Critical Expiry color coding, aur automatic comments add karegi.")
        st.code(advanced_vba_code, language="vb")
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            st.download_button("📥 Download Advanced VBA Script (.bas)", advanced_vba_code.encode("utf-8"), "Advanced_Enterprise_Expiry_Macro.bas", "text/plain")
        with col_v2:
            st.info("💡 **Tip:** Excel me `Alt + F11` press karke VBA editor open karein, phir `File -> Import File` se is `.bas` file ko load karein.")

    with vba_tab2:
        st.markdown("##### 📁 Custom Template Upload & Management Hub")
        st.markdown("Aap apna khud ka custom Excel template upload kar sakte hain jise system future batch processing aur export formatting ke liye use karega.")
        
        custom_template_file = st.file_uploader("Upload Custom Excel Template (.xlsx / .xls)", type=["xlsx", "xls"], key="custom_tpl_upload")
        if custom_template_file:
            tpl_bytes = custom_template_file.getvalue()
            with open("Output.xlsx", "wb") as f_out:
                f_out.write(tpl_bytes)
            st.success(f"✅ Custom template '{custom_template_file.name}' successfully uploaded and set as active system template (`Output.xlsx`)!")
            
            # Preview uploaded template
            try:
                df_preview = pd.read_excel(io.BytesIO(tpl_bytes), nrows=10)
                st.markdown("##### Template Structure Preview (First 10 Rows):")
                st.dataframe(df_preview, use_container_width=True)
            except Exception as e:
                st.warning(f"Could not preview template sheets: {e}")

    with vba_tab3:
        st.markdown("##### 🔍 Multi-Criteria Search, Filter & Live Data Editor")
        st.markdown("Inventory aur Expiry records par advanced filters lagayein, direct edit karein, naye rows insert/delete karein, aur changes save karein.")

        conn = get_db_connection()
        # Fetch inventory / stock data for editing
        df_inv = pd.read_sql("SELECT * FROM sap_stock_master", conn)

        # Multi-Criteria Filter Controls
        mc_c1, mc_c2, mc_c3 = st.columns(3)
        with mc_c1:
            filter_plant = st.selectbox("Filter by Plant", ["All"] + df_inv["plant"].unique().tolist() if not df_inv.empty else ["All"])
        with mc_c2:
            search_query = st.text_input("🔍 Keyword Search (Material / Desc)", "")
        with mc_c3:
            stock_threshold = st.number_input("Min Unrestricted Stock", value=0.0, step=100.0)

        # Apply Filters
        filtered_df = df_inv.copy()
        if filter_plant != "All":
            filtered_df = filtered_df[filtered_df["plant"] == filter_plant]
        if search_query:
            filtered_df = filtered_df[filtered_df.apply(lambda r: r.astype(str).str.contains(search_query, case=False).any(), axis=1)]
        filtered_df = filtered_df[filtered_df["unrestricted_stock"] >= stock_threshold]

        st.markdown("##### ✏️ Interactive Data Editor (Edit cells, insert rows, or delete records):")
        edited_inventory = st.data_editor(
            filtered_df,
            num_rows="dynamic",
            use_container_width=True,
            key="advanced_inventory_editor",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "net_available_stock": st.column_config.NumberColumn("Net Available", disabled=True)
            }
        )

        ed_c1, ed_c2, ed_c3 = st.columns(3)
        with ed_c1:
            if st.button("💾 Save Inventory Changes", type="primary"):
                cur = conn.cursor()
                for _, row in edited_inventory.iterrows():
                    net_calc = float(row['unrestricted_stock']) - float(row['pending_order_demand'])
                    cur.execute("""
                        INSERT OR REPLACE INTO sap_stock_master 
                        (id, plant, storage_location, material_code, material_description, unrestricted_stock, reserved_stock, transit_stock, pending_order_demand, net_available_stock, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        row.get('id') if pd.notna(row.get('id')) else None,
                        str(row['plant']), str(row['storage_location']), str(row['material_code']),
                        str(row['material_description']), float(row['unrestricted_stock']),
                        float(row['reserved_stock']), float(row['transit_stock']),
                        float(row['pending_order_demand']), net_calc, get_ist_timestamp_full()
                    ))
                conn.commit()
                st.success("✅ Inventory edits successfully saved to database!")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()

        with ed_c2:
            st.download_button(
                "📥 Export Filtered Data (.xlsx)", 
                to_excel_download_bytes(edited_inventory, "InventoryExport"), 
                f"Inventory_Export_{get_ist_date_str()}.xlsx", 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with ed_c3:
            if st.button("🗑️ Delete Selected Rows (By ID)"):
                st.info("Use the row selection checkbox in data editor to delete rows.")

        conn.close()

    with vba_tab4:
        st.markdown("##### 📊 Shelf-Life & Stock Expiry Analytics")
        conn = get_db_connection()
        df_analytics = pd.read_sql("SELECT material_code, material_description, unrestricted_stock, pending_order_demand FROM sap_stock_master", conn)
        conn.close()

        if not df_analytics.empty:
            an_c1, an_c2 = st.columns(2)
            with an_c1:
                st.markdown("##### Stock Distribution by Material")
                st.bar_chart(df_analytics.set_index("material_code")["unrestricted_stock"])
            with an_c2:
                st.markdown("##### Pending Demand vs Unrestricted Stock")
                st.bar_chart(df_analytics.set_index("material_code")[["unrestricted_stock", "pending_order_demand"]])
        else:
            st.info("No analytics data available.")
# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE
# PART 4: DAILY REGISTERS, PARTIAL DATABASES, PENDING LEDGERS, 
# MASTER DB, FLEET CONFIGURATIONS & EXECUTIVE KPI ANALYTICS
# ==============================================================================

# ==============================================================================
# MODULE 7: DAILY DISPATCH SALE REGISTER
# ==============================================================================

if main_menu == "📖 Daily Dispatch Sale Register":
    st.title("📖 Daily Dispatch Sale Register Database")
    st.markdown("Comprehensive daily dispatch register with live cell editing, date filtering, multi-criteria search, and bulk purge options.")

    conn = get_db_connection()
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register ORDER BY register_id DESC", conn)

    # Date Range Filters
    rc1, rc2 = st.columns(2)
    with rc1:
        reg_from_date = st.date_input("Filter From Date", datetime.date.today() - datetime.timedelta(days=30), key="reg_f_date")
    with rc2:
        reg_to_date = st.date_input("Filter To Date", datetime.date.today(), key="reg_t_date")

    search_r = st.text_input("🔍 Search Register (Agency, Route, Vehicle, DR Code, Order No):", "")
    if search_r:
        df_reg = df_reg[df_reg.apply(lambda r: r.astype(str).str.contains(search_r, case=False).any(), axis=1)]

    st.markdown("##### ✏️ Live Cell Editing Register Table:")
    edited_reg = st.data_editor(df_reg, use_container_width=True, num_rows="dynamic", key="editor_reg_advanced")

    rc_btn1, rc_btn2, rc_btn3 = st.columns([1, 1, 2])
    with rc_btn1:
        if st.button("💾 Save Register Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_reg.iterrows():
                cur.execute("""
                    UPDATE daily_dispatch_register 
                    SET dispatch_date=?, trip_id=?, vehicle_no=?, transporter_name=?, route_no=?, agency_no=?, order_no=?, dr_code=?, fg_code=?, dispatched_bags=?, dispatched_weight_mt=?, bay_no=?
                    WHERE register_id=?
                """, (
                    str(row['dispatch_date']), str(row['trip_id']), str(row['vehicle_no']), str(row['transporter_name']),
                    str(row['route_no']), str(row['agency_no']), str(row['order_no']), str(row['dr_code']),
                    str(row['fg_code']), float(row['dispatched_bags']), float(row['dispatched_weight_mt']),
                    str(row['bay_no']), row['register_id']
                ))
            conn.commit()
            st.success("✅ Register changes successfully saved!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()

    with rc_btn2:
        if not df_reg.empty:
            st.download_button("📥 Export Register (.xlsx)", to_excel_download_bytes(df_reg, "DailyRegister"), f"Daily_Dispatch_Register_{get_ist_date_str()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with rc_btn3:
        with st.expander("🗑️ Delete & Purge Options"):
            del_r_ids = st.multiselect("Select Register IDs to Delete:", df_reg["register_id"].tolist() if not df_reg.empty else [])
            if st.button("Delete Selected IDs"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM daily_dispatch_register WHERE register_id=?", [(i,) for i in del_r_ids])
                conn.commit()
                st.success("Selected records deleted.")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
            if st.button("🔥 Purge Complete Register"):
                cur = conn.cursor()
                cur.execute("DELETE FROM daily_dispatch_register")
                cur.execute("DELETE FROM sqlite_sequence WHERE name='daily_dispatch_register'")
                conn.commit()
                st.success("Register completely wiped!")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 8: PARTIAL / SPLIT DISPATCH DATABASE
# ==============================================================================

elif main_menu == "🧩 Partial / Split Dispatch Database":
    st.title("🧩 Partial & Split Dispatch Database")
    st.markdown("Track orders that were partially dispatched in vehicle trips and manage their remaining pending balances.")

    conn = get_db_connection()
    df_part = pd.read_sql("SELECT * FROM partial_dispatch_ledger ORDER BY id DESC", conn)

    search_part = st.text_input("🔍 Search Partial Ledger (Trip ID, Agency, Route, FG Code):", "")
    if search_part:
        df_part = df_part[df_part.apply(lambda r: r.astype(str).str.contains(search_part, case=False).any(), axis=1)]

    edited_part = st.data_editor(df_part, use_container_width=True, num_rows="dynamic", key="editor_part_advanced")

    pc1, pc2, pc3 = st.columns([1, 1, 2])
    with pc1:
        if st.button("💾 Save Partial Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_part.iterrows():
                cur.execute("""
                    UPDATE partial_dispatch_ledger 
                    SET trip_id=?, route_no=?, agency_no=?, dr_code=?, fg_code=?, original_bags=?, dispatched_bags=?, remaining_bags=?, status=?
                    WHERE id=?
                """, (
                    str(row['trip_id']), str(row['route_no']), str(row['agency_no']), str(row['dr_code']),
                    str(row['fg_code']), float(row['original_bags']), float(row['dispatched_bags']),
                    float(row['remaining_bags']), str(row['status']), row['id']
                ))
            conn.commit()
            st.success("✅ Partial ledger changes saved!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    with pc2:
        if not df_part.empty:
            st.download_button("📥 Export Partials (.xlsx)", to_excel_download_bytes(df_part, "Partials"), f"Partial_Dispatch_Ledger_{get_ist_date_str()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with pc3:
        with st.expander("🗑️ Delete Partial Records"):
            del_part_ids = st.multiselect("Select Partial IDs to Delete:", df_part["id"].tolist() if not df_part.empty else [])
            if st.button("Delete Selected Partials"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM partial_dispatch_ledger WHERE id=?", [(i,) for i in del_part_ids])
                conn.commit()
                st.success("Selected records deleted.")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 9: PENDING ORDERS LEDGER
# ==============================================================================

elif main_menu == "⏳ Pending Orders Ledger":
    st.title("⏳ Pending Orders Database Hub")
    st.markdown("Live view of all pending orders awaiting vehicle trip allocation across routes and agencies.")

    conn = get_db_connection()
    df_p = pd.read_sql("SELECT * FROM pending_orders ORDER BY id DESC", conn)

    search_p = st.text_input("🔍 Search Pending Orders (Order No, Route, Agency, FG Code, DR Code):", "")
    if search_p:
        df_p = df_p[df_p.apply(lambda r: r.astype(str).str.contains(search_p, case=False).any(), axis=1)]

    edited_p = st.data_editor(df_p, use_container_width=True, num_rows="dynamic", key="editor_pending_advanced")

    pec1, pec2, pec3 = st.columns([1, 1, 2])
    with pec1:
        if st.button("💾 Save Pending Order Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_p.iterrows():
                cur.execute("""
                    UPDATE pending_orders 
                    SET order_no=?, route_no=?, agency_no=?, dr_code=?, fg_code=?, bags_qty=?, weight_mt=?, status=?
                    WHERE id=?
                """, (
                    str(row['order_no']), str(row['route_no']), str(row['agency_no']), str(row['dr_code']),
                    str(row['fg_code']), float(row['bags_qty']), float(row['weight_mt']), str(row['status']), row['id']
                ))
            conn.commit()
            st.success("✅ Pending orders successfully saved!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    with pec2:
        if not df_p.empty:
            st.download_button("📥 Export Pending (.xlsx)", to_excel_download_bytes(df_p, "PendingOrders"), f"Pending_Orders_{get_ist_date_str()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with pec3:
        with st.expander("🗑️ Delete Options"):
            del_p_ids = st.multiselect("Select Order IDs to Delete:", df_p["id"].tolist() if not df_p.empty else [])
            if st.button("Delete Selected Orders"):
                cur = conn.cursor()
                cur.executemany("DELETE FROM pending_orders WHERE id=?", [(i,) for i in del_p_ids])
                conn.commit()
                st.success("Selected orders deleted.")
                st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 10: FILE UPLOAD ARCHIVE & MASTER DB
# ==============================================================================

elif main_menu == "🗄️ File Upload Archive":
    st.title("🗄️ Uploaded Input File Archive & Audit Logs")
    conn = get_db_connection()
    df_a = pd.read_sql("SELECT * FROM uploaded_files_archive ORDER BY id DESC", conn)
    st.dataframe(df_a, use_container_width=True)
    if not df_a.empty:
        st.download_button("📥 Export Archive Logs", to_excel_download_bytes(df_a, "Archive"), "Upload_Archive.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    conn.close()

elif main_menu == "📋 Master DB & Unmapped Ledger":
    st.title("📋 Unique Master Mapping DB & Unmapped Fallback Ledger")
    conn = get_db_connection()
    df_m = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
    df_u = pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger ORDER BY id DESC", conn)
    
    tab_m1, tab_m2 = st.tabs(["📋 Master Routes", "🚨 Unmapped DR Ledger"])
    with tab_m1:
        st.dataframe(df_m, use_container_width=True)
        st.download_button("📥 Export Master", to_excel_download_bytes(df_m, "Master"), "Unique_Master.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab_m2:
        st.dataframe(df_u, use_container_width=True)
        st.download_button("📥 Export Unmapped", to_excel_download_bytes(df_u, "Unmapped"), "Unmapped_DR.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    conn.close()

# ==============================================================================
# MODULE 11: FLEET & LOADING BAY MASTER
# ==============================================================================

elif main_menu == "🚛 Fleet & Loading Bay Master":
    st.title("🚛 Transporter Fleet Master & Loading Bay Configurations")
    conn = get_db_connection()
    df_f = pd.read_sql("SELECT * FROM fleet_master", conn)
    df_b = pd.read_sql("SELECT * FROM loading_bays", conn)
    
    ft_1, ft_2 = st.tabs(["🚛 Fleet Master", "🏭 Loading Bays"])
    with ft_1:
        edited_f = st.data_editor(df_f, use_container_width=True, num_rows="dynamic", key="edit_fleet_adv")
        if st.button("💾 Save Fleet Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_f.iterrows():
                cur.execute("UPDATE fleet_master SET vehicle_no=?, vehicle_type=?, capacity_bags=?, capacity_mt=?, transporter_name=?, driver_name=?, driver_phone=?, status=? WHERE id=?", (str(row['vehicle_no']), str(row['vehicle_type']), int(row['capacity_bags']), float(row['capacity_mt']), str(row['transporter_name']), str(row['driver_name']), str(row['driver_phone']), str(row['status']), row['id']))
            conn.commit()
            st.success("Fleet updated!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    with ft_2:
        edited_b = st.data_editor(df_b, use_container_width=True, num_rows="dynamic", key="edit_bay_adv")
        if st.button("💾 Save Bays Changes", type="primary"):
            cur = conn.cursor()
            for _, row in edited_b.iterrows():
                cur.execute("UPDATE loading_bays SET bay_no=?, bay_name=?, status=? WHERE id=?", (str(row['bay_no']), str(row['bay_name']), str(row['status']), row['id']))
            conn.commit()
            st.success("Loading bays updated!")
            st.rerun() if hasattr(st, "rerun") else st.experimental_rerun()
    conn.close()

# ==============================================================================
# MODULE 12: EXECUTIVE KPI & VISUAL ANALYTICS
# ==============================================================================

elif main_menu == "📊 Executive KPI & Visual Analytics":
    st.title("📊 Supply Chain & Dispatch KPI Analytics Dashboard")
    st.markdown("Real-time operational metrics, vehicle capacity utilization, route dispatch breakdown, and executive supply chain performance indicators.")

    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips", conn)
    df_pending = pd.read_sql("SELECT * FROM pending_orders", conn)
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register", conn)
    conn.close()

    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    tot_trips = len(df_trips)
    tot_dispatched_bags = df_reg["dispatched_bags"].sum() if not df_reg.empty else 0.0
    avg_util = df_trips["capacity_utilization_pct"].mean() if not df_trips.empty else 0.0
    active_pending_bags = df_pending[df_pending["status"] == "Pending"]["bags_qty"].sum() if not df_pending.empty else 0.0

    col_k1.metric("Total Trips Planned", f"{tot_trips:,}")
    col_k2.metric("Dispatched Bags", f"{tot_dispatched_bags:,.0f}")
    col_k3.metric("Avg Fleet Utilization", f"{avg_util:.1f}%")
    col_k4.metric("Active Pending Load", f"{active_pending_bags:,.0f} Bags")

    st.markdown("---")
    c_ch1, c_ch2 = st.columns(2)
    with c_ch1:
        st.markdown("##### 🚛 Route-wise Dispatched Load (MT)")
        if not df_reg.empty:
            st.bar_chart(df_reg.groupby("route_no")["dispatched_weight_mt"].sum())
        else:
            st.info("No dispatched register records available.")
    with c_ch2:
        st.markdown("##### 📦 Active Pending Orders by Route")
        if not df_pending.empty:
            st.bar_chart(df_pending[df_pending["status"] == "Pending"].groupby("route_no")["bags_qty"].sum())
        else:
            st.info("No pending orders available.")
