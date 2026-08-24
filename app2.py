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
    page_title="Enterprise Sales Order & Dispatch Automation Hub", 
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

# --- SQLite Database Initialization with Enhanced Dispatch Master & Input File Ledgers ---
def init_db():
    conn = sqlite3.connect("sales_history.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS history_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            files_count INTEGER,
            total_qty REAL,
            status TEXT
        )
    """)
    # Enhanced Master Dispatch & Unique Routes Table containing complete details as requested
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS unique_routes_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            agency_name TEXT,
            dr_code TEXT,
            fg_code TEXT,
            quantity REAL,
            vehicle_no TEXT,
            mobile_no TEXT,
            created_at TEXT
        )
    """)
    # Raw Input Files Archive & Database Ledger with Delete/Wipe support
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS input_files_archive_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            file_blob BLOB,
            file_size_kb REAL,
            uploaded_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS output_files_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            file_type TEXT,
            file_data BLOB,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            route_no TEXT,
            agency_no TEXT,
            dr_code TEXT,
            created_at TEXT
        )
    """)
    cursor.execute("""
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
    cursor.execute("""
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
    conn.commit()
    conn.close()

init_db()

# --- Session State Defaults ---
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

# Professional CSS Injection
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
        .stButton>button:hover {{ background-color: {t['btn_hover']} !important; }}
        
        button[kind="primary"] {{
            background-color: {t['primary']} !important;
            color: #ffffff !important;
        }}
        button[kind="primary"] p {{ color: #ffffff !important; }}
        
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

# --- TOP COLLAPSIBLE ERP CONTROL PANEL & INSTANT THEME SELECTOR ---
with st.expander("⚙️ Enterprise Control Panel, Theme Engine & System Settings", expanded=False):
    st.subheader("🎨 Theme Engine (8 Professional Themes)")
    def on_theme_change():
        st.session_state.selected_theme = st.session_state.theme_selectbox

    st.selectbox(
        "Select Interface Theme", 
        list(THEMES.keys()), 
        key="theme_selectbox",
        index=list(THEMES.keys()).index(st.session_state.selected_theme),
        on_change=on_theme_change,
        label_visibility="collapsed"
    )
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

# Mapping assignments
default_fg_code = st.session_state.fg_code
col_mapping_input = st.session_state.col_map
agency_fg_override = st.session_state.agency_override
default_fallback_route = st.session_state.route
email_user = st.session_state.email_user
email_pass = st.session_state.email_pass
recipient_email = st.session_state.recipient
whatsapp_num = st.session_state.whatsapp

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

st.title(f"💼 Enterprise Sales Order & Dispatch Automation Hub ({st.session_state.selected_theme})")
st.markdown("Upload multiple **Inbound Demand Files** to process orders, archive input files in database, auto-lookup missing DRs, log unmapped entries, and track complete dispatch details (Route, Agency, DR, FG, Quantity, Vehicle No, Mobile No).")
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

if st.button("🚀 Process Batch Orders & Update Master Dispatch & Input DB", type="primary"):
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
        input_files_archive_records = []
        unmapped_records_to_insert = []
        output_files_to_store = []
        traceability_records = []
        discrepancy_records = []
        
        with st.spinner("⚡ Processing files, archiving inputs in DB, extracting Dispatch details & updating Master... Please wait."):
            try:
                try:
                    with open("Output.xlsx", "rb") as f:
                        template_bytes = f.read()
                except FileNotFoundError:
                    st.error("❌ 'Output.xlsx' template file repository mein nahi mili. Kripya template file ko upload karein.")
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
                    file_size_kb = len(file_bytes) / 1024.0
                    
                    # Archive Input File in Database
                    input_files_archive_records.append((short_filename, file_bytes, file_size_kb, batch_ts))

                    df_input = pd.read_excel(io.BytesIO(file_bytes), header=None)

                    # 1. Find FG Row & Col
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

                    # 2. Total/Sum Column Detection
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

                    # 3. Route Number Detection
                    route_num = default_fallback_route if default_fallback_route != "" else "22"
                    ignore_list = ["RT", "DR", "RT DR", "ROUTE", "SALES PERSON", "CONTACT NO:", "MATERIAL CODE"]
                    for r in range(fg_row):
                        for c in range(min(total_col, 30)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if upper_val in ignore_list or any(upper_val.startswith(p) for p in ["PC", "MS", "M", "GM", "DP", "SKU", "FG"]):
                                continue
                            if cell_val != "" and 1 <= len(cell_val) <= 3 and any(char.isdigit() for char in cell_val):
                                route_num = cell_val
                                break
                        if route_num != (default_fallback_route if default_fallback_route != "" else "22"):
                            break

                    if default_fallback_route != "" and default_fallback_route != "22":
                        route_num = default_fallback_route

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 4. Agency Column & Name Column Detection
                    agency_col = -1
                    agency_name_col = -1
                    mobile_col = -1
                    vehicle_col = -1

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

                    # Scan for Name, Mobile, and Vehicle columns in header or data rows
                    for cSearch in range(fg_col):
                        for rCheck in range(max(0, fg_row - 5), fg_row + 1):
                            hdr_val = str(df_input.iloc[rCheck, cSearch]).strip().upper()
                            if "NAME" in hdr_val and agency_name_col == -1:
                                agency_name_col = cSearch
                            elif "MOBILE" in hdr_val or "CONTACT" in hdr_val:
                                mobile_col = cSearch
                            elif "VEHICLE" in hdr_val or "PB" in hdr_val:
                                vehicle_col = cSearch

                    if agency_name_col == -1 and agency_col > 0:
                        agency_name_col = agency_col + 1

                    # DR Code Column Detection
                    dr_code_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        matched_count = 0
                        for offset in range(1, min(4, df_input.shape[0] - fg_row)):
                            v = str(df_input.iloc[fg_row + offset, cSearch]).strip().upper()
                            if re.match(r'^DR\d+', v):
                                matched_count += 1
                        if matched_count > 0:
                            dr_code_col = cSearch
                            break

                    # Valid FG Columns Collection
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

                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        if pd.isna(agency) or str(agency).strip() in ["", "nan", "None"]:
                            continue
                        
                        agency_str = str(agency).replace('.0','').strip()
                        if not agency_str.isdigit() or not (1 <= len(agency_str) <= 5):
                            st.session_state.skipped_rows_log.append({
                                "File Name": short_filename,
                                "Row Index": r + 1,
                                "Agency Value": str(agency),
                                "Reason": "Invalid or Non-numeric Agency Number"
                            })
                            total_skipped_rows += 1
                            continue

                        agency_val = int(agency_str)
                        
                        # Extract Agency Name, Mobile, Vehicle if present
                        ag_name = str(df_input.iloc[r, agency_name_col]).strip() if (agency_name_col >= 0 and agency_name_col < df_input.shape[1]) else "N/A"
                        mob_no = str(df_input.iloc[r, mobile_col]).strip() if (mobile_col >= 0 and mobile_col < df_input.shape[1]) else "N/A"
                        veh_no = str(df_input.iloc[r, vehicle_col]).strip() if (vehicle_col >= 0 and vehicle_col < df_input.shape[1]) else "N/A"

                        row_has_items = False
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
                                        row_has_items = True
                                        row_total_qty += qty_val
                                        file_input_qty += qty_val
                                        valid_row_quantities.append((c, fg_code, qty_val))
                                except ValueError:
                                    pass

                        if not row_has_items:
                            total_skipped_rows += 1
                            continue

                        # DR Code Detection
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

                        # Auto-lookup missing DR from Master DB
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
                            
                            st.session_state.unmapped_current_batch.append({
                                "File Name": short_filename,
                                "Route": str(route_num),
                                "Agency": agency_val,
                                "Status": "Generated via NEW_CUST (Missing DR)"
                            })

                        final_dr = clean_dr if has_dr_code else f"NEW_CUST_{agency_val}"

                        item_id = 10
                        for c, fg_code, qty_val in valid_row_quantities:
                            cleaned_fg = str(fg_code).strip()
                            upper_fg = cleaned_fg.upper()
                            
                            if (agency_val, c) in agency_col_override_map:
                                current_fg = agency_col_override_map[(agency_val, c)]
                            elif upper_fg.startswith("FG"):
                                current_fg = cleaned_fg
                            else:
                                current_fg = direct_col_mapping.get(c, default_fg_code)
                            
                            total_input_qty += qty_val
                            total_gen_qty += qty_val
                            
                            # Insert into Master Dispatch DB record with all details
                            db_records_to_insert.append((
                                short_filename, 
                                str(route_num), 
                                str(agency_val), 
                                ag_name, 
                                str(final_dr).upper(), 
                                current_fg, 
                                qty_val, 
                                veh_no, 
                                mob_no, 
                                ist_now.strftime("%Y-%m-%d %H:%M:%S")
                            ))
                            
                            file_comparison_rows.append({
                                "File Name": short_filename,
                                "Status": "Valid DR" if has_dr_code else "Missing DR",
                                "Agency": agency_val,
                                "Agency Name": ag_name,
                                "DR Code": final_dr,
                                "FG Code": current_fg,
                                "Input Qty": qty_val,
                                "Generated Qty": qty_val,
                                "Vehicle No": veh_no,
                                "Mobile No": mob_no
                            })

                            target_ws = ws_valid if has_dr_code else ws_missing
                            current_r = valid_row if has_dr_code else missing_row
                            order_num = valid_order_num if has_dr_code else missing_order_num
                            
                            target_ws.cell(row=current_r, column=2, value=order_num)
                            target_ws.cell(row=current_r, column=3, value="OR")
                            target_ws.cell(row=current_r, column=4, value="SO20")
                            target_ws.cell(row=current_r, column=5, value=10)
                            target_ws.cell(row=current_r, column=6, value=20)
                            target_ws.cell(row=current_r, column=7, value=final_dr)
                            target_ws.cell(row=current_r, column=8, value=final_dr)
                            target_ws.cell(row=current_r, column=9, value=f"RT-{route_num}-{agency_val}-{today_date}")
                            target_ws.cell(row=current_r, column=10, value=today_date)
                            target_ws.cell(row=current_r, column=11, value=today_date)
                            target_ws.cell(row=current_r, column=15, value=item_id)
                            target_ws.cell(row=current_r, column=16, value=current_fg)
                            target_ws.cell(row=current_r, column=19, value=qty_val)
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
                        st.session_state.processed_files.append({"name": short_filename + " (Valid DR)", "data": buf_valid.getvalue(), "filename": out_fname})
                        output_files_to_store.append((out_fname, "Valid DR", buf_valid.getvalue(), ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname, "Valid DR", 1, batch_ts))

                    if missing_items_created > 0:
                        buf_missing = io.BytesIO()
                        wb_missing.save(buf_missing)
                        buf_missing.seek(0)
                        out_fname_miss = safe_route_num + "_" + today_date + "_" + timestamp + "_Missing_DR.xlsx"
                        st.session_state.processed_files.append({"name": short_filename + " (Missing DR)", "data": buf_missing.getvalue(), "filename": out_fname_miss})
                        output_files_to_store.append((out_fname_miss, "Missing DR", buf_missing.getvalue(), ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                        traceability_records.append((batch_ts, short_filename, file_bytes, file_input_qty, out_fname_miss, "Missing DR", 1, batch_ts))

                    if file_comparison_rows:
                        st.session_state.comparison_summary.append(pd.DataFrame(file_comparison_rows))

                # --- Write to Databases ---
                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                
                # Insert Archived Input Files
                cursor.executemany("""
                    INSERT INTO input_files_archive_ledger (file_name, file_blob, file_size_kb, uploaded_at)
                    VALUES (?, ?, ?, ?)
                """, input_files_archive_records)

                # Insert into Master Dispatch DB
                cursor.executemany("""
                    INSERT INTO unique_routes_master (file_name, route_no, agency_no, agency_name, dr_code, fg_code, quantity, vehicle_no, mobile_no, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, db_records_to_insert)

                cursor.executemany("""
                    INSERT INTO unmapped_missing_dr_ledger (file_name, route_no, agency_no, dr_code, created_at)
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

                st.success("✅ Batch Processed, Input Files Archived & Master Dispatch Database Updated Successfully!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya pehle demand files upload karein!")

# KPI Summary & Visual Dashboards
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
    st.markdown("### 📥 Bulk Download & Actions")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])
    
    col_zip, col_email = st.columns(2)
    with col_zip:
        st.download_button(
            label="📦 Download All Output Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Batch_Orders_{get_ist_now().strftime('%Y-%m-%d')}.zip",
            mime="application/zip",
            key="zip_download"
        )

# --- ALL DATABASES MANAGEMENT PANEL WITH DELETE & WIPE OPTIONS ---
st.markdown("---")
with st.expander("🗄️ Manage Databases: Master Dispatch, Archived Input Files, Unmapped & Outputs (Delete & Wipe Options)", expanded=True):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_master = pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn)
        df_inputs_arch = pd.read_sql("SELECT id, file_name, file_size_kb, uploaded_at FROM input_files_archive_ledger ORDER BY id DESC", conn)
        df_unmapped = pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger ORDER BY id DESC", conn)
        df_outputs = pd.read_sql("SELECT id, file_name, file_type, created_at FROM output_files_ledger ORDER BY id DESC", conn)
        conn.close()
        
        tab_m1, tab_m2, tab_m3, tab_m4 = st.tabs([
            "🚚 Master Dispatch Database", 
            "📂 Archived Input Files Database", 
            "🚨 Unmapped Missing DR", 
            "📦 Output Files Ledger"
        ])
        
        # --- TAB 1: MASTER DISPATCH DATABASE MANAGEMENT ---
        with tab_m1:
            st.markdown("#### 🚚 Master Dispatch & Route Details (Route, Agency, DR, FG, Vehicle, Mobile)")
            if not df_master.empty:
                db_search = st.text_input("🔍 Search Master Dispatch DB", "", key="master_search")
                filtered_master = df_master
                if db_search:
                    q = db_search.lower()
                    filtered_master = df_master[
                        df_master['file_name'].astype(str).str.lower().str.contains(q) |
                        df_master['route_no'].astype(str).str.lower().str.contains(q) |
                        df_master['agency_name'].astype(str).str.lower().str.contains(q) |
                        df_master['dr_code'].astype(str).str.lower().str.contains(q) |
                        df_master['vehicle_no'].astype(str).str.lower().str.contains(q) |
                        df_master['mobile_no'].astype(str).str.lower().str.contains(q)
                    ]
                st.dataframe(filtered_master, use_container_width=True)
                
                st.markdown("##### 🗑️ Delete & Wipe Tools (Master Dispatch)")
                d1, d2 = st.columns(2)
                with d1:
                    master_del_id = st.number_input("Enter Master Record ID to Delete", min_value=1, step=1, key="m_del_id")
                    if st.button("🗑️ Delete Master Record & Reset ID"):
                        conn = sqlite3.connect("sales_history.db")
                        cur = conn.cursor()
                        cur.execute("DELETE FROM unique_routes_master WHERE id = ?", (master_del_id,))
                        cur.execute("DELETE FROM sqlite_sequence WHERE name='unique_routes_master'")
                        conn.commit()
                        conn.close()
                        st.success(f"✅ Master Record ID {master_del_id} deleted & ID sequence reset!")
                        st.rerun()
                with d2:
                    if st.button("🚨 Wipe Entire Master Dispatch DB & Reset IDs"):
                        conn = sqlite3.connect("sales_history.db")
                        cur = conn.cursor()
                        cur.execute("DELETE FROM unique_routes_master")
                        cur.execute("DELETE FROM sqlite_sequence WHERE name='unique_routes_master'")
                        conn.commit()
                        conn.close()
                        st.success("✅ Master Dispatch DB wiped & IDs reset!")
                        st.rerun()

                master_buf = io.BytesIO()
                df_master.to_excel(master_buf, index=False, sheet_name="Master Dispatch")
                master_buf.seek(0)
                st.download_button(
                    label="📥 Export Master Dispatch DB to Excel (.xlsx)",
                    data=master_buf.getvalue(),
                    file_name=f"Master_Dispatch_Database_{get_ist_now().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_master_dispatch"
                )
            else:
                st.info("No master dispatch records found yet.")

        # --- TAB 2: ARCHIVED INPUT FILES DATABASE MANAGEMENT ---
        with tab_m2:
            st.markdown("#### 📂 Archived Input Files Database (Upload History & Blobs)")
            if not df_inputs_arch.empty:
                st.dataframe(df_inputs_arch, use_container_width=True)
                
                st.markdown("##### 📥 Download or 🗑️ Delete Archived Input Files")
                in_col1, in_col2 = st.columns(2)
                with in_col1:
                    sel_in_id = st.number_input("Enter Archived Input File ID", min_value=1, step=1, key="sel_in_id")
                    if st.button("📥 Download Original Input File"):
                        conn = sqlite3.connect("sales_history.db")
                        cur = conn.cursor()
                        cur.execute("SELECT file_name, file_blob FROM input_files_archive_ledger WHERE id = ?", (sel_in_id,))
                        res = cur.fetchone()
                        conn.close()
                        if res:
                            st.download_button(
                                label=f"💾 Save '{res[0]}'",
                                data=res[1],
                                file_name=res[0],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_in_blob_{sel_in_id}"
                            )
                        else:
                            st.warning("⚠️ Invalid ID.")
                
                with in_col2:
                    del_in_id = st.number_input("Enter Input File ID to Delete", min_value=1, step=1, key="del_in_id")
                    if st.button("🗑️ Delete Input File Record & Reset ID"):
                        conn = sqlite3.connect("sales_history.db")
                        cur = conn.cursor()
                        cur.execute("DELETE FROM input_files_archive_ledger WHERE id = ?", (del_in_id,))
                        cur.execute("DELETE FROM sqlite_sequence WHERE name='input_files_archive_ledger'")
                        conn.commit()
                        conn.close()
                        st.success(f"✅ Input File Record ID {del_in_id} deleted & IDs reset!")
                        st.rerun()
                    if st.button("🚨 Wipe Entire Input Archive DB & Reset IDs"):
                        conn = sqlite3.connect("sales_history.db")
                        cur = conn.cursor()
                        cur.execute("DELETE FROM input_files_archive_ledger")
                        cur.execute("DELETE FROM sqlite_sequence WHERE name='input_files_archive_ledger'")
                        conn.commit()
                        conn.close()
                        st.success("✅ Input Files Database wiped & IDs reset!")
                        st.rerun()
            else:
                st.info("No input files archived in database yet.")

        # --- TAB 3: UNMAPPED LEDGER ---
        with tab_m3:
            st.markdown("#### 🚨 Unmapped Missing DR Ledger")
            if not df_unmapped.empty:
                st.dataframe(df_unmapped, use_container_width=True)
                if st.button("🚨 Wipe Unmapped Ledger & Reset IDs", key="wipe_unmap"):
                    conn = sqlite3.connect("sales_history.db")
                    cur = conn.cursor()
                    cur.execute("DELETE FROM unmapped_missing_dr_ledger")
                    cur.execute("DELETE FROM sqlite_sequence WHERE name='unmapped_missing_dr_ledger'")
                    conn.commit()
                    conn.close()
                    st.success("✅ Unmapped Ledger wiped & IDs reset!")
                    st.rerun()
            else:
                st.info("No unmapped missing DR records logged.")

        # --- TAB 4: OUTPUT FILES LEDGER ---
        with tab_m4:
            st.markdown("#### 📦 Archived Output Files Ledger")
            if not df_outputs.empty:
                st.dataframe(df_outputs, use_container_width=True)
                if st.button("🚨 Wipe Output Ledger & Reset IDs", key="wipe_out"):
                    conn = sqlite3.connect("sales_history.db")
                    cur = conn.cursor()
                    cur.execute("DELETE FROM output_files_ledger")
                    cur.execute("DELETE FROM sqlite_sequence WHERE name='output_files_ledger'")
                    conn.commit()
                    conn.close()
                    st.success("✅ Output Files Ledger wiped & IDs reset!")
                    st.rerun()
            else:
                st.info("No output files archived.")

    except Exception as e:
        st.error(f"Error managing databases: {str(e)}")

# Summary table view
if st.session_state.comparison_summary:
    st.markdown("---")
    st.markdown("### 📋 Dispatch Summary Table")
    combined_df = pd.concat(st.session_state.comparison_summary, ignore_index=True)
    st.dataframe(combined_df, use_container_width=True)
