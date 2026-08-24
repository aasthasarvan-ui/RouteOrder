# ==============================================================================
# ENTERPRISE LOGISTICS, DISPATCH ENGINE & SALES AUTOMATION SUITE (FIXED PARSER)
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
# SECTION 1: GLOBAL SETUP & TIMEZONE (IST)
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

def get_ist_timestamp_full():
    return get_ist_now().strftime("%Y-%m-%d %H:%M:%S")

def get_ist_file_suffix():
    return get_ist_now().strftime("%H%M%S")

DB_NAME = "enterprise_logistics_sales_hub.db"

# ==============================================================================
# SECTION 2: THEME PALETTES
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
if "fg_code" not in st.session_state:
    st.session_state.fg_code = "FG500014"
if "col_map" not in st.session_state:
    st.session_state.col_map = "36:FG500014AJ\n37:FG500014AK"
if "agency_override" not in st.session_state:
    st.session_state.agency_override = "101:36:FG500014N01\n101:37:FG500014N02"
if "route" not in st.session_state:
    st.session_state.route = "22"
if "whatsapp_num" not in st.session_state:
    st.session_state.whatsapp_num = "919876543210"
if "processed_files" not in st.session_state:
    st.session_state.processed_files = []
if "skipped_rows_log" not in st.session_state:
    st.session_state.skipped_rows_log = []
if "kpi_data" not in st.session_state:
    st.session_state.kpi_data = {"input_qty": 0.0, "gen_qty": 0.0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}

t = THEMES.get(st.session_state.selected_theme, THEMES["💼 Classic Enterprise Navy"])

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: {t['bg']}; color: {t['text']}; font-family: 'Segoe UI', Tahoma, sans-serif; }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{ color: {t['text']} !important; }}
        input, textarea, select {{ background-color: {t['input_bg']} !important; color: {t['input_text']} !important; border: 1px solid {t['border']} !important; border-radius: 4px; }}
        .stButton>button {{ width: 100%; height: 38px; background-color: {t['btn_bg']} !important; color: #ffffff !important; font-size: 13px !important; font-weight: 600 !important; border-radius: 4px; border: 1px solid {t['border']}; }}
        button[kind="primary"] {{ background-color: {t['primary']} !important; color: #ffffff !important; }}
        div[data-testid="stExpander"] {{ background-color: {t['card_bg']}; border: 1px solid {t['border']}; border-radius: 6px; }}
        div[data-testid="stDataFrame"] {{ border: 1px solid {t['border']}; border-radius: 6px; background-color: {t['card_bg']}; }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================================================================
# SECTION 3: DATABASE ARCHITECTURE
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
            uploaded_at TEXT,
            UNIQUE(order_no, route_no, agency_no, fg_code, bags_qty)
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

    # Seed Default Fleet
    cur.execute("SELECT COUNT(*) FROM fleet_master")
    if cur.fetchone()[0] == 0:
        cur.executemany("""
            INSERT INTO fleet_master (vehicle_no, vehicle_type, capacity_bags, capacity_mt, transporter_name, driver_name, driver_phone, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, [
            ('PB-10-AZ-1122', '10 Wheeler Truck', 400, 20.0, 'National Logistics', 'Gurpreet Singh', '9876543210', 'Available'),
            ('PB-08-BX-4455', '12 Wheeler Multi-Axle', 600, 30.0, 'Speedway Cargo', 'Baljit Sharma', '9812345678', 'Available')
        ])

    cur.execute("SELECT COUNT(*) FROM loading_bays")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO loading_bays (bay_no, bay_name, status) VALUES (?, ?, ?)", [
            ('BAY-01', 'North Plant Main Gate', 'Open'),
            ('BAY-02', 'Storage Silo Bay 2', 'Open')
        ])

    conn.commit()
    conn.close()

init_all_enterprise_databases()

def to_excel_download_bytes(df: pd.DataFrame, sheet_name="DataSheet") -> bytes:
    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output_stream.getvalue()

# ==============================================================================
# SECTION 4: NAVIGATION & SIDEBAR
# ==============================================================================

with st.sidebar:
    st.image("https://img.icons8.com/color/96/delivery-truck.png", width=55)
    st.title("Logistics Master Suite")

    main_menu = st.radio(
        "Navigation",
        [
            "⚡ Inbound Demand & Sales Order Engine",
            "🚚 Route Dispatch Trip Planner",
            "📋 Loading Slips & Active Trips",
            "📖 Daily Dispatch Sale Register",
            "⏳ Pending Orders Ledger",
            "🗄️ File Upload Archive",
            "📋 Master DB & Unmapped Ledger",
            "🚛 Fleet & Loading Bay Master"
        ]
    )

# ==============================================================================
# MODULE 1: INBOUND DEMAND & SALES ORDER AUTOMATION (FIXED EXTRACTION ENGINE)
# ==============================================================================

if main_menu == "⚡ Inbound Demand & Sales Order Engine":
    st.title("⚡ Enterprise Inbound Demand & Sales Order Processing Engine")
    st.markdown("Upload demand workbooks to parse SKUs, match DR codes, and generate Output templates.")

    with st.expander("⚙️ SKU & Route Settings", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.session_state.fg_code = st.text_input("Default Fallback FG Code", value=st.session_state.fg_code)
            st.session_state.route = st.text_input("Default Route Fallback", value=st.session_state.route)
        with c2:
            st.session_state.col_map = st.text_area("Column Index Mapping (Col:FG)", value=st.session_state.col_map, height=80)
        with c3:
            st.session_state.agency_override = st.text_area("Agency SKU Overrides (Agency:Col:FG)", value=st.session_state.agency_override, height=80)

    col_map_dict = {}
    for line in st.session_state.col_map.split("\n"):
        if ":" in line:
            p = line.split(":")
            if p[0].strip().isdigit():
                col_map_dict[int(p[0].strip())] = p[1].strip()

    agency_override_dict = {}
    for line in st.session_state.agency_override.split("\n"):
        p = line.split(":")
        if len(p) == 3 and p[0].strip().isdigit() and p[1].strip().isdigit():
            agency_override_dict[(int(p[0].strip()), int(p[1].strip()))] = p[2].strip()

    uploaded_files = st.file_uploader("Upload Inbound Demand Excel Workbooks", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files and st.button("🚀 Process Batch Orders & Ingest to Pending Database", type="primary"):
        st.session_state.processed_files = []
        st.session_state.skipped_rows_log = []

        total_in_qty = 0.0
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

            f_bytes = up_file.getvalue()
            try:
                df_input = pd.read_excel(io.BytesIO(f_bytes), header=None)
            except Exception as e:
                st.error(f"Error reading '{short_fname}': {str(e)}")
                continue

            # 1. Broad FG Search (Exact or Partial)
            fg_row, fg_col = -1, -1
            for r in range(min(df_input.shape[0], 25)):
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

            # 2. Total Column Detection
            total_col = df_input.shape[1]
            for c_s in range(fg_col, df_input.shape[1]):
                val = str(df_input.iloc[fg_row, c_s]).strip().upper()
                if any(kw in val for kw in ["TOTAL", "SUM", "NET", "TTL", "GR. TOTAL"]):
                    total_col = c_s
                    break

            # 3. Route Number Fallback & Detection
            route_num = st.session_state.route if st.session_state.route != "" else "22"
            for r in range(fg_row):
                for c in range(min(total_col, 20)):
                    val = str(df_input.iloc[r, c]).replace('.0', '').strip()
                    if val != "" and 1 <= len(val) <= 4 and any(ch.isdigit() for ch in val):
                        if not any(prefix in val.upper() for prefix in ["DR", "FG", "OR", "SO", "RT-"]):
                            route_num = val
                            break

            resolved_route = "".join(ch for ch in str(route_num) if ch.isalnum() or ch in ('-', '_'))

            # 4. Agency & DR Column
            agency_col = -1
            for c_s in range(fg_col - 1, -1, -1):
                col_samples = df_input.iloc[fg_row + 1: fg_row + 10, c_s].dropna().astype(str).str.replace(r'\.0$', '', regex=True)
                if col_samples.str.isdigit().sum() >= 2:
                    agency_col = c_s
                    break
            if agency_col == -1:
                agency_col = fg_col - 1 if fg_col > 0 else 0

            dr_code_col = -1
            for c_s in range(fg_col - 1, -1, -1):
                col_samples = df_input.iloc[fg_row + 1: fg_row + 10, c_s].dropna().astype(str).str.upper()
                if col_samples.str.startswith("DR").sum() >= 1:
                    dr_code_col = c_s
                    break

            valid_cols = [(c, str(df_input.iloc[fg_row, c]).strip()) for c in range(fg_col, total_col)]

            # Output Excel Workbook Init
            wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes)) if template_bytes else openpyxl.Workbook()
            ws_valid = wb_valid["Order Data"] if "Order Data" in wb_valid.sheetnames else wb_valid.active
            wb_missing = openpyxl.load_workbook(io.BytesIO(template_bytes)) if template_bytes else openpyxl.Workbook()
            ws_missing = wb_missing["Order Data"] if "Order Data" in wb_missing.sheetnames else wb_missing.active

            valid_r_idx, missing_r_idx = 6, 6
            valid_order_no, missing_order_no = 1, 1
            valid_items_cnt, missing_items_cnt = 0, 0
            agency_counts_valid, agency_counts_missing = {}, {}

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
                    cur.execute(
                        "SELECT dr_code FROM unique_routes_master WHERE route_no=? AND agency_no=?",
                        (resolved_route, str(agency_val))
                    )
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
                    if pd.notna(q_val):
                        try:
                            f_qty = float(q_val)
                            if f_qty > 0:
                                current_fg = fg_val if fg_val.startswith("FG") else col_map_dict.get(c_idx, st.session_state.fg_code)
                                if (agency_val, c_idx) in agency_override_dict:
                                    current_fg = agency_override_dict[(agency_val, c_idx)]

                                pending_records_to_insert.append((
                                    short_fname, f"ORD-{agency_val}-{r}", resolved_route, str(agency_val),
                                    clean_dr, current_fg, f_qty, round(f_qty * 0.05, 2), ref_code, "Pending", batch_ts
                                ))

                                total_in_qty += f_qty
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
                INSERT OR IGNORE INTO pending_orders (source_file, order_no, route_no, agency_no, dr_code, fg_code, bags_qty, weight_mt, order_ref, status, uploaded_at)
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

        conn.commit()
        conn.close()

        st.session_state.kpi_data = {
            "input_qty": total_in_qty, "gen_qty": total_in_qty, "valid_count": total_valid,
            "missing_count": total_missing, "skipped_count": total_skipped
        }

        if len(st.session_state.processed_files) > 0:
            st.success(f"🎉 Batch execution completed! Created {len(st.session_state.processed_files)} output workbooks & saved {len(pending_records_to_insert)} pending orders.")
        else:
            st.warning("⚠️ Koi valid demand row extract nahi ho saki. Kripya check karein ki Excel me valid positive quantities hain.")

    if st.session_state.processed_files:
        st.markdown("---")
        st.subheader("📥 Generated Output Files & Downloads")
        for idx_f, f_itm in enumerate(st.session_state.processed_files):
            st.download_button(
                f"📥 Download {f_itm['name']}",
                f_itm["data"],
                f_itm["filename"],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_btn_{idx_f}"
            )

# ==============================================================================
# MODULE 2: ROUTE DISPATCH TRIP PLANNER
# ==============================================================================

elif main_menu == "🚚 Route Dispatch Trip Planner":
    st.title("🚚 Route Dispatch Planning & Vehicle Allocation")
    conn = get_db_connection()
    df_pending = pd.read_sql("SELECT * FROM pending_orders WHERE status='Pending'", conn)

    if df_pending.empty:
        st.info("ℹ️ No pending orders found. Please upload demand files.")
    else:
        route_summary = df_pending.groupby("route_no").agg({"agency_no": "nunique", "bags_qty": "sum", "weight_mt": "sum"}).reset_index()
        st.dataframe(route_summary, use_container_width=True)

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            sel_route = st.selectbox("Select Route:", route_summary["route_no"].tolist())
            avail_fleet = pd.read_sql("SELECT * FROM fleet_master WHERE status='Available'", conn)
            avail_bays = pd.read_sql("SELECT * FROM loading_bays WHERE status='Open'", conn)
            fleet_opts = [f"{r['vehicle_no']} | {r['vehicle_type']} ({r['capacity_bags']} Bags)" for _, r in avail_fleet.iterrows()]
            sel_vehicle = st.selectbox("Assign Vehicle:", fleet_opts if fleet_opts else ["No Vehicles"])
            sel_bay = st.selectbox("Assign Bay:", [f"{r['bay_no']} - {r['bay_name']}" for _, r in avail_bays.iterrows()])
        with col_p2:
            route_df = df_pending[df_pending["route_no"] == str(sel_route)]
            agencies = route_df["agency_no"].unique().tolist()
            sel_agencies = st.multiselect("Select Agencies:", agencies, default=agencies)
            trip_df = route_df[route_df["agency_no"].isin(sel_agencies)]
            st.metric("Total Bags Selected", f"{trip_df['bags_qty'].sum():,.0f}")

        if st.button("🚀 Confirm Trip", type="primary"):
            if sel_vehicle != "No Vehicles" and not trip_df.empty:
                v_num = sel_vehicle.split(" | ")[0]
                now_ist = get_ist_now()
                trip_id = f"TRIP-{sel_route}-{now_ist.strftime('%Y%m%d%H%M%S')}"
                v_info = avail_fleet[avail_fleet["vehicle_no"] == v_num].iloc[0]
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO trip_loading_slips (trip_id, trip_date, route_no, vehicle_no, transporter_name, driver_name, driver_phone, loading_bay, total_bags, total_weight_mt, capacity_utilization_pct, status, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Planned', ?)
                """, (trip_id, now_ist.strftime("%Y-%m-%d"), str(sel_route), v_num, v_info["transporter_name"], v_info["driver_name"], v_info["driver_phone"], sel_bay.split(" - ")[0], trip_df["bags_qty"].sum(), trip_df["weight_mt"].sum(), 100, now_ist.strftime("%Y-%m-%d %H:%M:%S")))
                for seq, (_, r_val) in enumerate(trip_df.iterrows(), 1):
                    cur.execute("""
                        INSERT INTO trip_order_items (trip_id, order_no, agency_no, route_no, dr_code, fg_code, allocated_bags, allocated_weight_mt, delivery_seq)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (trip_id, r_val["order_no"], r_val["agency_no"], r_val["route_no"], r_val["dr_code"], r_val["fg_code"], r_val["bags_qty"], r_val["weight_mt"], seq))
                    cur.execute("UPDATE pending_orders SET status='Assigned' WHERE id=?", (r_val["id"],))
                cur.execute("UPDATE fleet_master SET status='Assigned to Trip' WHERE vehicle_no=?", (v_num,))
                conn.commit()
                st.success("Trip Created!")
                st.rerun()
    conn.close()

# ==============================================================================
# MODULE 3: ACTIVE TRIPS & LOADING SLIPS
# ==============================================================================

elif main_menu == "📋 Loading Slips & Active Trips":
    st.title("📋 Active Trips & Loading Slips")
    conn = get_db_connection()
    df_trips = pd.read_sql("SELECT * FROM trip_loading_slips ORDER BY created_at DESC", conn)
    st.dataframe(df_trips, use_container_width=True)
    if not df_trips.empty:
        sel_trip = st.selectbox("Select Trip ID:", df_trips["trip_id"].tolist())
        trip_row = df_trips[df_trips["trip_id"] == sel_trip].iloc[0]
        items_df = pd.read_sql("SELECT * FROM trip_order_items WHERE trip_id=?", conn, params=(sel_trip,))
        st.dataframe(items_df, use_container_width=True)
        if trip_row["status"] != "Dispatched" and st.button("🏁 Mark Dispatched", type="primary"):
            cur = conn.cursor()
            for _, it in items_df.iterrows():
                cur.execute("""
                    INSERT INTO daily_dispatch_register (dispatch_date, trip_id, vehicle_no, transporter_name, route_no, agency_no, order_no, dr_code, fg_code, dispatched_bags, dispatched_weight_mt, bay_no, dispatched_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (trip_row["trip_date"], trip_row["trip_id"], trip_row["vehicle_no"], trip_row["transporter_name"], trip_row["route_no"], it["agency_no"], it["order_no"], it["dr_code"], it["fg_code"], it["allocated_bags"], it["allocated_weight_mt"], trip_row["loading_bay"], get_ist_timestamp_full()))
            cur.execute("UPDATE trip_loading_slips SET status='Dispatched' WHERE trip_id=?", (sel_trip,))
            cur.execute("UPDATE fleet_master SET status='Available' WHERE vehicle_no=?", (trip_row["vehicle_no"],))
            conn.commit()
            st.success("Dispatched!")
            st.rerun()
    conn.close()

# ==============================================================================
# MODULE 4: DAILY DISPATCH SALE REGISTER
# ==============================================================================

elif main_menu == "📖 Daily Dispatch Sale Register":
    st.title("📖 Daily Dispatch Sale Register")
    conn = get_db_connection()
    df_reg = pd.read_sql("SELECT * FROM daily_dispatch_register ORDER BY register_id DESC", conn)
    st.dataframe(df_reg, use_container_width=True)
    if not df_reg.empty:
        st.download_button("📥 Export to Excel", to_excel_download_bytes(df_reg, "SaleRegister"), "Daily_Dispatch_Sale_Register.xlsx")
    conn.close()

# ==============================================================================
# MODULE 5: PENDING ORDERS LEDGER
# ==============================================================================

elif main_menu == "⏳ Pending Orders Ledger":
    st.title("⏳ Pending Orders Ledger")
    conn = get_db_connection()
    df_p = pd.read_sql("SELECT * FROM pending_orders ORDER BY id DESC", conn)
    st.dataframe(df_p, use_container_width=True)
    if not df_p.empty:
        st.download_button("📥 Export Pending to Excel", to_excel_download_bytes(df_p, "Pending"), "Pending_Orders.xlsx")
    conn.close()

# ==============================================================================
# MODULE 6: ARCHIVE, MASTER & FLEET
# ==============================================================================

elif main_menu == "🗄️ File Upload Archive":
    st.title("🗄️ Uploaded Input File Archive")
    conn = get_db_connection()
    df_a = pd.read_sql("SELECT * FROM uploaded_files_archive ORDER BY id DESC", conn)
    st.dataframe(df_a, use_container_width=True)
    conn.close()

elif main_menu == "📋 Master DB & Unmapped Ledger":
    st.title("📋 Master DB & Unmapped Ledger")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM unique_routes_master ORDER BY id DESC", conn), use_container_width=True)
    conn.close()

elif main_menu == "🚛 Fleet & Loading Bay Master":
    st.title("🚛 Fleet & Loading Bay Master")
    conn = get_db_connection()
    st.dataframe(pd.read_sql("SELECT * FROM fleet_master", conn), use_container_width=True)
    conn.close()
