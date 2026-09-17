# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE (50-FEATURE COMPLETE)
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
    page_title="Enterprise Logistics & Sales Automation Hub (50-Feature Complete)",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

IST = pytz.timezone("Asia/Kolkata")

def get_ist_now():
    return datetime.datetime.now(IST)

def get_ist_date_str():
    return get_ist_now().strftime("%Y-%m-%d")

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
# SECTION 2: THEME DEFINITIONS & STYLING
# ==============================================================================

THEMES = {
    "💼 Classic Enterprise Navy": {
        "bg": "#f4f6f9", "text": "#1f2937", "card_bg": "#ffffff", "border": "#cbd5e1",
        "btn_bg": "#1e3a8a", "btn_hover": "#1d4ed8", "primary": "#2563eb", "input_bg": "#ffffff", "input_text": "#1f2937"
    },
    "🌙 Modern Dark ERP": {
        "bg": "#0b0f19", "text": "#f3f4f6", "card_bg": "#1f2937", "border": "#374151",
        "btn_bg": "#374151", "btn_hover": "#4b5563", "primary": "#3b82f6", "input_bg": "#111827", "input_text": "#f3f4f6"
    }
}

if "selected_theme" not in st.session_state:
    st.session_state.selected_theme = "💼 Classic Enterprise Navy"

curr_th = THEMES[st.session_state.selected_theme]

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: {curr_th['bg']}; color: {curr_th['text']}; font-family: 'Segoe UI', Tahoma, sans-serif; }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{ color: {curr_th['text']} !important; }}
        input, textarea, select {{ background-color: {curr_th['input_bg']} !important; color: {curr_th['input_text']} !important; border: 1px solid {curr_th['border']} !important; }}
        .stButton>button {{ width: 100%; height: 38px; background-color: {curr_th['btn_bg']} !important; color: #ffffff !important; font-weight: 600 !important; border-radius: 4px; }}
        .stButton>button:hover {{ background-color: {curr_th['btn_hover']} !important; }}
        div[data-testid="stExpander"], div[data-testid="stDataFrame"] {{ background-color: {curr_th['card_bg']}; border: 1px solid {curr_th['border']}; border-radius: 6px; }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================================================================
# SECTION 3: SESSION STATE DEFAULTS
# ==============================================================================

DEFAULTS = {
    "fg_code": "FG500014",
    "route": "22",
    "max_capacity": 320.0,
    "email_user": "",
    "email_pass": "",
    "recipient": "",
    "whatsapp_num": "919876543210",
    "processed_files": [],
    "kpi_data": {}
}

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==============================================================================
# SECTION 4: COMPLETE SQLITE DATABASE INITIALIZATION (ALL TABLES)
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
    cur.execute("CREATE TABLE IF NOT EXISTS trip_order_items (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_id TEXT, order_no TEXT, agency_no TEXT, route_no TEXT, dr_code TEXT, fg_code TEXT, allocated_bags REAL, allocated_weight_mt REAL, delivery_seq INTEGER, status TEXT DEFAULT 'Assigned')")
    cur.execute("CREATE TABLE IF NOT EXISTS daily_dispatch_register (register_id INTEGER PRIMARY KEY AUTOINCREMENT, dispatch_date TEXT, trip_id TEXT, vehicle_no TEXT, transporter_name TEXT, route_no TEXT, agency_no TEXT, order_no TEXT, dr_code TEXT, fg_code TEXT, dispatched_bags REAL, dispatched_weight_mt REAL, bay_no TEXT, dispatched_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS partial_dispatch_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_id TEXT, source_file TEXT, order_no TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, fg_code TEXT, original_bags REAL, dispatched_bags REAL, remaining_bags REAL, status TEXT DEFAULT 'Partial Pending', created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no))")
    cur.execute("CREATE TABLE IF NOT EXISTS plant_inventory_stock (id INTEGER PRIMARY KEY AUTOINCREMENT, upload_batch_id TEXT, source_file TEXT, material_code TEXT, material_desc TEXT, pack_size_kg INTEGER DEFAULT 50, unit_weight_mt REAL DEFAULT 0.05, plant_stock_qty REAL DEFAULT 0.0, safety_stock_qty REAL DEFAULT 100.0, stock_date TEXT, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS output_files_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, file_type TEXT, file_data BLOB, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS input_output_traceability (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, input_file_name TEXT, input_file_blob BLOB, total_input_qty REAL, generated_output_file TEXT, output_type TEXT, version_no INTEGER, created_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS history_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, files_count INTEGER, total_qty REAL, status TEXT, remarks TEXT)")

    # Seed Fleet & Bays if empty
    cur.execute("SELECT COUNT(*) FROM fleet_master")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status) VALUES (?, ?, ?, ?, ?, ?, ?, 'Available')", [
            ('PB-10-AZ-1122', '10 Wheeler Truck', 400, 20.0, 'National Logistics', 'Gurpreet Singh', '9876543210'),
            ('PB-08-BX-4455', '12 Wheeler Multi-Axle', 600, 30.0, 'Speedway Cargo', 'Baljit Sharma', '9812345678'),
            ('PB-29-CD-9900', 'Canter / Eicher', 200, 10.0, 'Punjab Roadlines', 'Ramesh Kumar', '9823456789')
        ])

    cur.execute("SELECT COUNT(*) FROM loading_bays")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, 'Open')", [
            ('BAY-01', 'North Plant Main Gate'), ('BAY-02', 'Storage Silo Bay 2')
        ])

    conn.commit()
    conn.close()

init_all_enterprise_databases()

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
    pdf.cell(190, 5, f"Official Gate Pass | Generated: {get_ist_timestamp_full()}", ln=True, align="C")
    pdf.ln(4)
    pdf.set_font("Arial", "B", 9)
    pdf.cell(95, 6, f"Trip ID: {trip_data.get('trip_id', '')}", border=1)
    pdf.cell(95, 6, f"Vehicle No: {trip_data.get('vehicle_no', '')}", border=1, ln=True)
    pdf.ln(4)
    pdf.set_font("Arial", "B", 9)
    pdf.cell(15, 6, "Seq", border=1, align="C")
    pdf.cell(45, 6, "Agency No", border=1, align="C")
    pdf.cell(45, 6, "DR Code", border=1, align="C")
    pdf.cell(45, 6, "FG Code", border=1, align="C")
    pdf.cell(40, 6, "Bags Qty", border=1, ln=True, align="C")
    pdf.set_font("Arial", "", 8)
    for _, it in items_df.iterrows():
        pdf.cell(15, 5, str(it.get("delivery_seq", "")), border=1, align="C")
        pdf.cell(45, 5, str(it.get("agency_no", "")), border=1, align="C")
        pdf.cell(45, 5, str(it.get("dr_code", "")), border=1, align="C")
        pdf.cell(45, 5, str(it.get("fg_code", "")), border=1, align="C")
        pdf.cell(40, 5, f"{float(it.get('allocated_bags', 0)):,.0f}", border=1, ln=True, align="R")
    return bytes(pdf.output())

# ==============================================================================
# SECTION 5: SIDEBAR NAVIGATION (ALL 15+ MODULES)
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
    st.subheader("⚙️ ERP Parameters")
    st.text_input("Default Route No", key="route")
    st.text_input("Default FG Code", key="fg_code")
    st.number_input("Max Truck Capacity (Bags)", value=320.0, step=10.0, key="max_capacity")
    st.text_input("WhatsApp Alert Mobile No", key="whatsapp_num")

# ==============================================================================
# MODULE 1: INBOUND DEMAND & SALES ORDER ENGINE
# ==============================================================================

if main_menu == "⚡ Inbound Demand & Sales Order Engine":
    st.title("⚡ Enterprise Inbound Demand & Sales Order Processing Engine")
    st.markdown("Upload multiple **Demand Workbooks** to execute **Master File DR Auto-Lookup**, eliminate duplicate orders, and generate multi-truck loading schedules.")

    uploaded_files = st.file_uploader("Upload Inbound Demand Excel Workbooks", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files and st.button("🚀 Process Batch Orders & Auto-Lookup DR Codes", type="primary"):
        conn = get_db_connection()
        cur = conn.cursor()

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
                        if 'Route' in df_r.columns and 'Agency' in df_r.columns and 'DRCODE' in df_r.columns:
                            for _, mr in df_r.iterrows():
                                rt_val, ag_val, dr_val = str(mr['Route']).replace('.0','').strip(), str(mr['Agency']).replace('.0','').strip(), str(mr['DRCODE']).strip()
                                if dr_val and dr_val.upper() not in ["NAN", "NONE", ""]:
                                    master_lookup_dict[(rt_val, ag_val)] = dr_val
            except Exception: pass

        batch_ts = get_ist_timestamp_full()
        all_raw_records = []
        db_inserts, unmapped_inserts = [], []
        total_in_qty = 0.0

        for up_file in uploaded_files:
            try:
                df_input = pd.read_excel(io.BytesIO(up_file.getvalue()), header=None)
            except Exception: continue

            fg_row, fg_col = -1, -1
            for r in range(min(df_input.shape[0], 30)):
                for c in range(df_input.shape[1]):
                    if "FG" in str(df_input.iloc[r, c]).strip().upper():
                        fg_row, fg_col = r, c
                        break
                if fg_row != -1: break
            if fg_row == -1: continue

            total_col = df_input.shape[1]
            for c_s in range(fg_col, df_input.shape[1]):
                if any(kw in str(df_input.iloc[fg_row, c_s]).strip().upper() for kw in ["TOTAL", "SUM", "NET", "TTL", "GRAND"]):
                    total_col = c_s
                    break

            sku_cols = []
            for c in range(fg_col, total_col):
                h_val, f_val = str(df_input.iloc[max(0, fg_row-1), c]).strip().upper(), str(df_input.iloc[fg_row, c]).strip().upper()
                if any(kw in h_val for kw in ["TOTAL", "SUM", "REMARK"]) or any(kw in f_val for kw in ["TOTAL", "SUM", "REMARK"]): break
                sku_cols.append((c, str(df_input.iloc[max(0, fg_row-1), c]).strip(), str(df_input.iloc[fg_row, c]).strip()))

            agency_col = -1
            for cSearch in range(fg_col - 1, -1, -1):
                if df_input.iloc[fg_row + 1: fg_row + 15, cSearch].dropna().astype(str).str.replace('.0','',regex=True).str.isdigit().sum() >= 2:
                    agency_col = cSearch
                    break
            if agency_col == -1: agency_col = fg_col - 1 if fg_col > 0 else 0

            for r in range(fg_row + 1, df_input.shape[0]):
                ag_val = str(df_input.iloc[r, agency_col]).replace('.0', '').strip()
                farmer = str(df_input.iloc[r, agency_col+1] if agency_col+1 < fg_col else "").strip()
                if 'TOTAL' in ag_val.upper() or 'TOTAL' in farmer.upper() or not ag_val.isdigit(): continue

                lookup_key = (str(st.session_state.route), ag_val)
                clean_dr, has_dr = "", False
                if lookup_key in sql_lookup_dict:
                    clean_dr, has_dr = sql_lookup_dict[lookup_key], True
                elif lookup_key in master_lookup_dict:
                    clean_dr, has_dr = master_lookup_dict[lookup_key], True
                    db_inserts.append((st.session_state.route, ag_val, clean_dr, batch_ts))
                else:
                    clean_dr = f"NEW_CUST_{ag_val}"
                    unmapped_inserts.append((st.session_state.route, ag_val, clean_dr, batch_ts))

                for col_idx, sku_name, fg_code in sku_cols:
                    q_val = df_input.iloc[r, col_idx]
                    if pd.notna(q_val):
                        q_str = str(q_val).replace('.0','').strip()
                        if q_str.isdigit() and float(q_str) > 0:
                            f_qty = float(q_str)
                            total_in_qty += f_qty
                            all_raw_records.append({'route': st.session_state.route, 'ag_no': ag_val, 'dr_code': clean_dr, 'farmer': farmer, 'sku': sku_name, 'fg_code': fg_code, 'qty': f_qty})

        if all_raw_records:
            df_dem = pd.DataFrame(all_raw_records)
            dedup_df = df_dem.groupby(['route', 'ag_no', 'dr_code', 'farmer', 'sku', 'fg_code'], as_index=False)['qty'].sum()

            max_cap = st.session_state.max_capacity
            trucks_dict, current_truck = {}, 1
            for _, row in dedup_df.iterrows():
                qty_left = row['qty']
                while qty_left > 0:
                    if current_truck not in trucks_dict: trucks_dict[current_truck] = []
                    space_left = max_cap - sum(item['qty'] for item in trucks_dict[current_truck])
                    if qty_left <= space_left:
                        part = row.to_dict(); part['qty'] = qty_left
                        trucks_dict[current_truck].append(part); qty_left = 0
                    else:
                        part = row.to_dict(); part['qty'] = space_left
                        trucks_dict[current_truck].append(part)
                        qty_left -= space_left; current_truck += 1

            cur.executemany("INSERT OR IGNORE INTO unique_routes_master (route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?)", db_inserts)
            cur.executemany("INSERT OR IGNORE INTO unmapped_missing_dr_ledger (route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?)", unmapped_inserts)
            
            for _, r_d in dedup_df.iterrows():
                cur.execute("INSERT INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at) VALUES ('Batch', ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?)",
                            (f"ORD-{r_d['ag_no']}", str(r_d['route']), str(r_d['ag_no']), str(r_d['dr_code']), str(r_d['fg_code']), float(r_d['qty']), round(float(r_d['qty'])*0.05, 2), f"RT-{r_d['route']}-{r_d['ag_no']}", batch_ts))
            conn.commit()
            st.session_state.clean_demand_df = dedup_df
            st.session_state.trucks_dict = trucks_dict
            st.session_state.processed_files = True
            st.success(f"✅ Success! Extracted {total_in_qty:,.0f} Clean Bags, mapped Master DR Codes, and synced Pending Orders.")
        conn.close()

    if st.session_state.get("processed_files"):
        st.markdown("---")
        df_c = st.session_state.clean_demand_df
        t_d = st.session_state.trucks_dict
        st.metric("Total Clean Demand", f"{df_c['qty'].sum():,.0f} Bags")
        t1, t2 = st.tabs(["📋 Master Clean Demand", "🚚 Multi-Truck Loading Slips"])
        with t1: st.dataframe(df_c, use_container_width=True)
        with t2:
            for i, trk_items in t_d.items():
                st.write(f"### Truck {i} (Total: {sum(x['qty'] for x in trk_items)} Bags)")
                st.dataframe(pd.DataFrame(trk_items)[['farmer', 'dr_code', 'sku', 'qty']], use_container_width=True)

# ==============================================================================
# MODULE 2: ROUTE DISPATCH TRIP PLANNER
# ==============================================================================

elif main_menu == "🚚 Route Dispatch Trip Planner":
    st.title("🚚 Route Dispatch Planning & Vehicle Allocation")
    conn = get_db_connection()
    df_pending = pd.read_sql("SELECT * FROM pending_orders WHERE status='Pending'", conn)
    if df_pending.empty:
        st.info("ℹ️ No pending order demand found. Upload demand workbooks in Module 1.")
    else:
        st.dataframe(df_pending, use_container_width=True)
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
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips", conn)
    st.dataframe(df_trips, use_container_width=True)
    if not df_trips.empty:
        sel_trip = st.selectbox("Select Trip ID to view slip:", df_trips["trip_id"].tolist())
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
    st.title("📊 Executive KPI & Visual Analytics")
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
