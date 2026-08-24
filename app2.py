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
    page_title="Multi-Vehicle Enterprise Dispatch Plan Hub", 
    page_icon="🚚", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 8 Professional Enterprise Themes
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

# --- SQLite Database Initialization with Dispatch Plan Ledger ---
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
    # Master Dispatch & Unique Routes Table
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
    # Dedicated Dispatch Plan Database (Route, Date, Vehicle base)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS dispatch_plans_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dispatch_date TEXT,
            route_no TEXT,
            vehicle_no TEXT,
            agency_no TEXT,
            agency_name TEXT,
            dr_code TEXT,
            mobile_no TEXT,
            fg_code TEXT,
            quantity REAL,
            status TEXT,
            created_at TEXT
        )
    """)
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
    conn.commit()
    conn.close()

init_db()

# Session State Defaults
DEFAULTS = {
    "fg_code": "FG500014",
    "col_map": "36:FG500014AJ\n37:FG500014AK",
    "agency_override": "101:36:FG500014N01\n101:37:FG500014N02",
    "route": "22",
    "vehicle_no": "PB29AH2491",
    "selected_theme": "💼 Classic Enterprise Navy",
    "processed_files": [],
    "comparison_summary": [],
    "skipped_rows_log": [],
    "kpi_data": {"input_qty": 0, "gen_qty": 0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val

t = THEMES[st.session_state.selected_theme]

# Professional CSS Styling
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
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{ color: {t['text']} !important; }}
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
        }}
        .stButton>button p {{ color: #ffffff !important; }}
        .stButton>button:hover {{ background-color: {t['btn_hover']} !important; }}
        button[kind="primary"] {{ background-color: {t['primary']} !important; }}
        div[data-testid="stExpander"] {{
            background-color: {t['card_bg']};
            border: 1px solid {t['border']};
            border-radius: 4px;
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# --- TOP CONTROL PANEL & THEME SELECTOR ---
with st.expander("⚙️ Enterprise Dispatch & Vehicle Settings Hub", expanded=False):
    st.subheader("🎨 Theme Engine")
    def on_theme_change():
        st.session_state.selected_theme = st.session_state.theme_selectbox
    st.selectbox("Select Theme", list(THEMES.keys()), key="theme_selectbox", index=list(THEMES.keys()).index(st.session_state.selected_theme), on_change=on_theme_change, label_visibility="collapsed")
    
    st.markdown("---")
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        st.session_state.route = st.text_input("Default Route No", value=st.session_state.route)
    with col_c2:
        st.session_state.vehicle_no = st.text_input("Default Vehicle No (e.g. PB29AH2491)", value=st.session_state.vehicle_no)
    with col_c3:
        st.session_state.fg_code = st.text_input("Default FG Code", value=st.session_state.fg_code)

st.title(f"🚚 Multi-Vehicle Enterprise Dispatch & Screenshot-Style Plan Hub ({st.session_state.selected_theme})")
st.markdown("Upload demand files, assign specific **Vehicle Numbers**, **Routes**, and **Dispatch Dates**, and manage/print route-wise dispatch plans.")
st.markdown("---")

uploaded_inputs = st.file_uploader("Upload Inbound Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

# Dispatch Planning Form (Vehicle & Date Assignment)
with st.form("dispatch_plan_form"):
    st.subheader("🚛 Assign Vehicle & Dispatch Plan Parameters")
    f_col1, f_col2, f_col3 = st.columns(3)
    with f_col1:
        dispatch_date_input = st.date_input("Dispatch Date", get_ist_now().date())
    with f_col2:
        assigned_vehicle = st.text_input("Assigned Vehicle Number", value=st.session_state.vehicle_no)
    with f_col3:
        assigned_route = st.text_input("Route Number", value=st.session_state.route)

    process_btn = st.form_submit_button("🚀 Generate Vehicle Dispatch Plan & Save to Database", type="primary")

if process_btn:
    if uploaded_inputs:
        st.session_state.processed_files = []
        st.session_state.comparison_summary = []
        st.session_state.skipped_rows_log = []
        
        total_input_qty = 0
        total_gen_qty = 0
        total_valid_orders = 0
        total_skipped_rows = 0
        
        dispatch_records_to_insert = []
        input_files_archive_records = []
        output_files_to_store = []
        
        with st.spinner("⚡ Processing demand files and mapping vehicle dispatch schedules..."):
            try:
                try:
                    with open("Output.xlsx", "rb") as f:
                        template_bytes = f.read()
                except FileNotFoundError:
                    st.error("❌ 'Output.xlsx' template file repository mein nahi mili. Kripya upload karein.")
                    st.stop()
                
                ist_now = get_ist_now()
                date_str = dispatch_date_input.strftime("%Y-%m-%d")
                timestamp = ist_now.strftime("%H%M%S")
                batch_ts = ist_now.strftime("%Y-%m-%d %H:%M:%S")

                for uploaded_file in uploaded_inputs:
                    short_filename = uploaded_file.name
                    if short_filename.lower() == "output.xlsx": continue
                    
                    file_bytes = uploaded_file.getvalue()
                    input_files_archive_records.append((short_filename, file_bytes, len(file_bytes)/1024.0, batch_ts))
                    
                    df_input = pd.read_excel(io.BytesIO(file_bytes), header=None)

                    fg_row, fg_col = -1, -1
                    for r in range(df_input.shape[0]):
                        for c in range(df_input.shape[1]):
                            if "FG" in str(df_input.iloc[r, c]).strip().upper():
                                fg_row, fg_col = r, c
                                break
                        if fg_row != -1: break

                    if fg_row == -1: continue

                    total_col = df_input.shape[1]
                    for cSearch in range(fg_col, df_input.shape[1]):
                        if any(kw in str(df_input.iloc[r, cSearch]).strip().upper() for r in range(max(0, fg_row-5), fg_row+2) for kw in ["TOTAL", "SUM", "TTL"]):
                            total_col = cSearch
                            break

                    agency_col, agency_name_col, mobile_col = fg_col - 1, fg_col - 2, fg_col - 3
                    dr_code_col = -1
                    for cSearch in range(fg_col):
                        for rCheck in range(max(0, fg_row-5), fg_row+1):
                            val = str(df_input.iloc[rCheck, cSearch]).strip().upper()
                            if "DR" in val: dr_code_col = cSearch

                    valid_cols = []
                    for c in range(fg_col, total_col):
                        fg_code = str(df_input.iloc[fg_row, c]).strip()
                        if not any(kw in fg_code.upper() for kw in ["TOTAL", "SUM", "TTL"]):
                            valid_cols.append((c, fg_code))

                    wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_valid = wb_valid["Order Data"] if "Order Data" in wb_valid.sheetnames else wb_valid.active
                    
                    valid_row = 6
                    file_comp = []

                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        if pd.isna(agency) or str(agency).strip() in ["", "nan", "None"]: continue
                        
                        agency_str = str(agency).replace('.0','').strip()
                        if not agency_str.isdigit(): continue
                        agency_val = int(agency_str)
                        
                        ag_name = str(df_input.iloc[r, max(0, agency_name_col)]).strip() if agency_name_col >= 0 else f"Agency_{agency_val}"
                        mob_no = str(df_input.iloc[r, max(0, mobile_col)]).strip() if mobile_col >= 0 else "N/A"
                        
                        row_quantities = []
                        row_tot = 0
                        for c, fg_code in valid_cols:
                            qty = df_input.iloc[r, c]
                            if pd.notna(qty) and str(qty).strip() != "":
                                try:
                                    q_val = float(qty)
                                    if q_val > 0:
                                        row_tot += q_val
                                        total_input_qty += q_val
                                        total_gen_qty += q_val
                                        row_quantities.append((fg_code, q_val))
                                except: pass

                        if row_tot == 0: continue

                        # Extract or assign DR code
                        final_dr = f"DR{agency_val}"
                        if dr_code_col >= 0:
                            d_val = str(df_input.iloc[r, dr_code_col]).strip()
                            if "DR" in d_val.upper(): final_dr = d_val

                        for fg_code, q_val in row_quantities:
                            clean_fg = fg_code if str(fg_code).upper().startswith("FG") else st.session_state.fg_code
                            
                            # Insert into Dispatch Plan Database
                            dispatch_records_to_insert.append((
                                date_str,
                                str(assigned_route),
                                str(assigned_vehicle),
                                str(agency_val),
                                ag_name,
                                final_dr,
                                mob_no,
                                clean_fg,
                                q_val,
                                "Planned",
                                batch_ts
                            ))
                            
                            file_comp.append({
                                "Vehicle No": assigned_vehicle,
                                "Route": assigned_route,
                                "Date": date_str,
                                "Agency": agency_val,
                                "Name": ag_name,
                                "DR Code": final_dr,
                                "FG Code": clean_fg,
                                "Quantity": q_val,
                                "Mobile": mob_no
                            })

                    if file_comp:
                        st.session_state.comparison_summary.append(pd.DataFrame(file_comp))
                        
                # Save to SQLite Database
                conn = sqlite3.connect("sales_history.db")
                cur = conn.cursor()
                cur.executemany("""
                    INSERT INTO dispatch_plans_ledger (dispatch_date, route_no, vehicle_no, agency_no, agency_name, dr_code, mobile_no, fg_code, quantity, status, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, dispatch_records_to_insert)
                conn.commit()
                conn.close()

                st.success(f"✅ Dispatch Plan successfully generated for Vehicle **{assigned_vehicle}** (Route: {assigned_route}) & saved to Database!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya demand files upload karein!")

# --- VIEW, FILTER, DOWNLOAD & PRINT DISPATCH PLANS ---
st.markdown("---")
st.markdown("### 🖨️ Vehicle, Route & Date-based Dispatch Plan Reports")

try:
    conn = sqlite3.connect("sales_history.db")
    df_dispatch_all = pd.read_sql("SELECT * FROM dispatch_plans_ledger ORDER BY id DESC", conn)
    conn.close()
    
    if not df_dispatch_all.empty:
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            all_vehicles = ["All Vehicles"] + df_dispatch_all['vehicle_no'].dropna().unique().tolist()
            sel_vehicle = st.selectbox("Filter by Vehicle Number", all_vehicles)
        with f_col2:
            all_routes = ["All Routes"] + df_dispatch_all['route_no'].dropna().unique().tolist()
            sel_route = st.selectbox("Filter by Route No", all_routes)
        with f_col3:
            all_dates = ["All Dates"] + df_dispatch_all['dispatch_date'].dropna().unique().tolist()
            sel_date = st.selectbox("Filter by Dispatch Date", all_dates)

        filtered_df = df_dispatch_all.copy()
        if sel_vehicle != "All Vehicles":
            filtered_df = filtered_df[filtered_df['vehicle_no'] == sel_vehicle]
        if sel_route != "All Routes":
            filtered_df = filtered_df[filtered_df['route_no'] == sel_route]
        if sel_date != "All Dates":
            filtered_df = filtered_df[filtered_df['dispatch_date'] == sel_date]

        st.markdown(f"#### Showing {len(filtered_df)} Dispatch Records")
        st.dataframe(filtered_df, use_container_width=True)

        # Download & Print Buttons for Filtered Plan
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            out_buf = io.BytesIO()
            filtered_df.to_excel(out_buf, index=False, sheet_name="Dispatch Plan")
            out_buf.seek(0)
            st.download_button(
                label="📥 Download Filtered Dispatch Plan (.xlsx)",
                data=out_buf.getvalue(),
                file_name=f"Dispatch_Plan_{sel_vehicle}_{sel_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with d_col2:
            print_html = """
            <div style="width:100%; margin-top:5px;">
                <button onclick="parent.window.print()" style="width:100%; height:38px; background:#1e3a8a; color:white; border:none; border-radius:4px; font-weight:600; cursor:pointer; display:flex; align-items:center; justify-content:center;">
                    🖨️ Print Dispatch Plan
                </button>
            </div>
            """
            components.html(print_html, height=50)

        # Database Management & Deletion Tools
        st.markdown("---")
        st.markdown("### 🗑️ Dispatch Database Management, Delete & Wipe Tools")
        del_c1, del_c2 = st.columns(2)
        with del_c1:
            rec_id_del = st.number_input("Enter Dispatch Record ID to Delete", min_value=1, step=1, key="disp_del")
            if st.button("🗑️ Delete Record & Reset IDs"):
                conn = sqlite3.connect("sales_history.db")
                cur = conn.cursor()
                cur.execute("DELETE FROM dispatch_plans_ledger WHERE id = ?", (rec_id_del,))
                cur.execute("DELETE FROM sqlite_sequence WHERE name='dispatch_plans_ledger'")
                conn.commit()
                conn.close()
                st.success(f"✅ Record ID {rec_id_del} deleted & ID sequence reset!")
                st.rerun()
        with del_c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚨 Wipe Entire Dispatch Plan DB & Reset IDs"):
                conn = sqlite3.connect("sales_history.db")
                cur = conn.cursor()
                cur.execute("DELETE FROM dispatch_plans_ledger")
                cur.execute("DELETE FROM sqlite_sequence WHERE name='dispatch_plans_ledger'")
                conn.commit()
                conn.close()
                st.success("✅ Dispatch Plans Database wiped & IDs reset!")
                st.rerun()

    else:
        st.info("No dispatch plans created yet. Upload demand files and assign vehicles above.")

except Exception as e:
    st.error(f"Error loading dispatch database: {str(e)}")
