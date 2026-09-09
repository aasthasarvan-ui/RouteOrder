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
from copy import copy
import os

# ==============================================================================
# SECTION 1: STREAMLIT PAGE CONFIGURATION & METADATA
# ==============================================================================
st.set_page_config(
    page_title="Enterprise Sales Order Automation Hub", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

master_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"

# ==============================================================================
# SECTION 2: 8 ENTERPRISE COLOR PALETTES & THEME DEFINITIONS
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
# SECTION 3: TIMEZONE CONFIGURATION (IST)
# ==============================================================================
IST = pytz.timezone('Asia/Kolkata')

def get_ist_now():
    return datetime.datetime.now(IST)

# ==============================================================================
# SECTION 4: DATABASE INTEGRITY VERIFICATION & INITIALIZATION
# ==============================================================================
def init_db():
    conn = sqlite3.connect("sales_history.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS history_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, files_count INTEGER, total_qty REAL, status TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS unique_routes_master (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no, dr_code))")
    cursor.execute("CREATE TABLE IF NOT EXISTS output_files_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, file_type TEXT, file_data BLOB, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS input_output_traceability (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, input_file_name TEXT, total_input_qty REAL, generated_output_file TEXT, output_type TEXT, created_at TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS discrepancy_audit_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_timestamp TEXT, file_name TEXT, agency_no TEXT, dr_code TEXT, fg_code TEXT, input_qty REAL, generated_qty REAL, difference REAL, logged_at TEXT)")
    conn.commit()
    conn.close()

init_db()

# ==============================================================================
# SECTION 5: SESSION STATE DEFAULTS
# ==============================================================================
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

# ==============================================================================
# SECTION 6: HIGH CONTRAST CSS INJECTION
# ==============================================================================
st.markdown(
    f"""
    <style>
        #GithubIcon {{ visibility: hidden !important; display: none !important; }}
        .stAppHeader {{ background-color: transparent !important; }}
        header[data-testid="stHeader"] {{ display: none !important; }}
        .stApp {{ background-color: {t['bg']}; color: {t['text']}; font-family: 'Segoe UI', sans-serif; }}
        h1, h2, h3, h4, h5, h6, p, span, label, .stMarkdown {{ color: {t['text']} !important; }}
        input, textarea, select {{ background-color: {t['input_bg']} !important; color: {t['input_text']} !important; border: 1px solid {t['border']} !important; }}
        .stButton>button {{ width: 100%; height: 38px; background-color: {t['btn_bg']} !important; color: #ffffff !important; font-weight: 600 !important; border-radius: 4px; }}
        .stButton>button:hover {{ background-color: {t['btn_hover']} !important; }}
    </style>
    """,
    unsafe_allow_html=True
)

# ==============================================================================
# SECTION 7: SIDEBAR - MANAGE MASTER DATA (EXACT COUNTIFS PRESERVATION)
# ==============================================================================
st.sidebar.header("🛠️ Manage Master Data")
action_choice = st.sidebar.radio("Choose Action", ["Add New Entry", "Delete Entry"])

if action_choice == "Add New Entry":
    with st.sidebar.form("add_master_form", clear_on_submit=True):
        st.subheader("➕ Add Master Record")
        new_route = st.text_input("Route Number")
        new_agency = st.text_input("Agency ID")
        new_bp = st.text_input("Business Partner Name")
        new_agency2 = st.text_input("Agency2 (Optional / Same as Agency)")
        new_drcode = st.text_input("DRCODE")

        submit_master = st.form_submit_button("Save to Master File")

        if submit_master:
            if not new_route or not new_agency or not new_drcode:
                st.sidebar.error("❌ Route, Agency, aur DRCODE bharna zaroori hai!")
            elif not os.path.exists(master_path):
                st.sidebar.error(f"❌ Master file ({master_path}) project folder mein nahi mili!")
            else:
                try:
                    r_raw = str(new_route).replace('.0', '').strip()
                    a_raw = str(new_agency).replace('.0', '').strip()
                    ag2_raw = str(new_agency2).strip().replace('.0', '') if new_agency2 else a_raw
                    bp_clean = str(new_bp).strip()
                    dr_clean = str(new_drcode).strip()

                    r_val = int(r_raw) if r_raw.isdigit() else r_raw
                    a_val = int(a_raw) if a_raw.isdigit() else a_raw
                    ag2_val = int(ag2_raw) if ag2_raw.isdigit() else ag2_raw

                    wb_m = openpyxl.load_workbook(master_path)
                    if "Master Data" in wb_m.sheetnames:
                        ws_m = wb_m["Master Data"]
                        is_duplicate = False
                        max_used_row = 2

                        for r in range(3, ws_m.max_row + 1):
                            cell_r = ws_m.cell(row=r, column=1).value
                            cell_a = ws_m.cell(row=r, column=2).value
                            if cell_r is not None or cell_a is not None:
                                max_used_row = r
                            if cell_r is not None and cell_a is not None:
                                if str(cell_r).replace('.0', '').strip() == str(r_val) and str(cell_a).replace('.0', '').strip() == str(a_val):
                                    is_duplicate = True
                                    break

                        if is_duplicate:
                            st.sidebar.error(f"⚠️ Duplicate Error: Route '{new_route}' aur Agency '{new_agency}' pehle se maujood hain!")
                        else:
                            target_row = max_used_row + 1
                            ws_m.cell(row=target_row, column=1, value=r_val)
                            ws_m.cell(row=target_row, column=2, value=a_val)
                            ws_m.cell(row=target_row, column=3, value=bp_clean)
                            ws_m.cell(row=target_row, column=4, value=ag2_val)
                            ws_m.cell(row=target_row, column=5, value=dr_clean)
                            ws_m.cell(row=target_row, column=6, value="YES")
                            ws_m.cell(row=target_row, column=7, value=f'=IF(F{target_row}="YES", A{target_row}&"_"&COUNTIFS($A$4:A{target_row}, A{target_row}, $F$4:F{target_row}, "YES"), "")')

                            for col in range(1, ws_m.max_column + 1):
                                prev_cell = ws_m.cell(row=max_used_row, column=col)
                                curr_cell = ws_m.cell(row=target_row, column=col)
                                if prev_cell.font: curr_cell.font = copy(prev_cell.font)
                                if prev_cell.border: curr_cell.border = copy(prev_cell.border)
                                if prev_cell.fill: curr_cell.fill = copy(prev_cell.fill)
                                if prev_cell.alignment: curr_cell.alignment = copy(prev_cell.alignment)
                                if col >= 8 and prev_cell.value and str(prev_cell.value).startswith('='):
                                    curr_cell.value = re.sub(r'(?<!\$)([A-Z]+)(\d+)', lambda m: f"{m.group(1)}{target_row}", str(prev_cell.value))

                            route_sheet_name = f"Route_{r_raw}"
                            if route_sheet_name in wb_m.sheetnames:
                                ws_r = wb_m[route_sheet_name]
                                max_r_row = ws_r.max_row
                                for r_chk in range(ws_r.max_row, 0, -1):
                                    if ws_r.cell(row=r_chk, column=1).value is not None or ws_r.cell(row=r_chk, column=2).value is not None:
                                        max_r_row = r_chk
                                        break
                                target_r_row = max_r_row + 1
                                ws_r.cell(row=target_r_row, column=1, value=r_val)
                                ws_r.cell(row=target_r_row, column=2, value=a_val)
                                ws_r.cell(row=target_r_row, column=3, value=bp_clean)
                                ws_r.cell(row=target_r_row, column=4, value=ag2_val)
                                ws_r.cell(row=target_r_row, column=5, value=dr_clean)
                                if max_r_row >= 3:
                                    for col in range(1, ws_r.max_column + 1):
                                        p_cell = ws_r.cell(row=max_r_row, column=col)
                                        c_cell = ws_r.cell(row=target_r_row, column=col)
                                        if p_cell.font: c_cell.font = copy(p_cell.font)
                                        if p_cell.border: c_cell.border = copy(p_cell.border)
                                        if p_cell.fill: c_cell.fill = copy(p_cell.fill)
                                        if p_cell.alignment: c_cell.alignment = copy(p_cell.alignment)
                                        if p_cell.value and str(p_cell.value).startswith('='):
                                            c_cell.value = re.sub(r'(?<!\$)([A-Z]+)(\d+)', lambda m: f"{m.group(1)}{target_r_row}", str(p_cell.value))

                            wb_m.save(master_path)
                            st.sidebar.success(f"🎉 Record successfully added to Master Data and {route_sheet_name}!")
                    else:
                        st.sidebar.error("❌ Master file mein 'Master Data' sheet nahi mili.")
                except Exception as ex:
                    st.sidebar.error(f"❌ Error saving master data: {ex}")

elif action_choice == "Delete Entry":
    st.sidebar.subheader("🗑️ Delete Master Record")
    if os.path.exists(master_path):
        try:
            temp_master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
            temp_master_df.columns = temp_master_df.columns.astype(str).str.strip()
            if 'Route' in temp_master_df.columns and 'Agency' in temp_master_df.columns:
                temp_master_df['Clean_Route'] = temp_master_df['Route'].astype(str).str.replace('.0', '', regex=False).str.strip()
                temp_master_df['Clean_Agency'] = temp_master_df['Agency'].astype(str).str.replace('.0', '', regex=False).str.strip()
                temp_master_df['Display_Label'] = "Route: " + temp_master_df['Clean_Route'] + " | Agency: " + temp_master_df['Clean_Agency'] + " | DRCODE: " + temp_master_df.get('DRCODE', '').astype(str)

                selected_to_delete = st.sidebar.selectbox("Select Record to Delete", temp_master_df['Display_Label'].tolist())
                if st.sidebar.button("Delete Selected Record"):
                    parts = selected_to_delete.split(" | ")
                    sel_route = parts[0].replace("Route: ", "").strip()
                    sel_agency = parts[1].replace("Agency: ", "").strip()
                    wb_m = openpyxl.load_workbook(master_path)
                    ws_m = wb_m["Master Data"]
                    row_to_delete = None
                    for r in range(3, ws_m.max_row + 1):
                        r_val = str(ws_m.cell(row=r, column=1).value).replace('.0', '').strip()
                        a_val = str(ws_m.cell(row=r, column=2).value).replace('.0', '').strip()
                        if r_val == sel_route and a_val == sel_agency:
                            row_to_delete = r
                            break
                    if row_to_delete:
                        ws_m.delete_rows(row_to_delete)
                        wb_m.save(master_path)
                        st.sidebar.success("🗑️ Record successfully delete ho gaya!")
                        st.rerun()
        except Exception as e:
            st.sidebar.error(f"❌ Error: {e}")

if os.path.exists(master_path):
    with open(master_path, "rb") as master_f:
        st.sidebar.download_button("📥 Download Updated Master File", data=master_f.read(), file_name="Business_Partners_Master_Updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==============================================================================
# SECTION 8: PRIMARY WORKFLOW: INBOUND DEMAND EXTRACTION & MASTER MAPPING
# ==============================================================================
st.title(f"💼 Enterprise Sales Order Automation Hub ({st.session_state.selected_theme})")
st.markdown("Upload multiple **Inbound Demand Files**. Background engine will auto-map DRCODEs from Master file (`Business_Partners_Master_Original_Keys_Restored.xlsx`), shift formulas, and process orders.")
st.markdown("---")

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")
master_file_input = st.file_uploader("Upload Master Route File (Optional - agar project folder mein nahi hai)", type=["xlsx", "xls"])

if uploaded_inputs:
    if st.button("🚀 Process Batch Orders & Update Master DB", type="primary"):
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

        output_files_to_store = []
        traceability_records = []
        unmapped_records_to_store = []

        with st.spinner("⚡ Loading Master File and Processing Batch Orders... Please wait."):
            try:
                master_df = None
                mapping_dict = {}

                if master_file_input is not None:
                    try:
                        master_df = pd.read_excel(master_file_input, sheet_name="Master Data", header=2)
                    except Exception:
                        pass
                elif os.path.exists(master_path):
                    try:
                        master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
                    except Exception:
                        pass

                if master_df is not None:
                    master_df.columns = master_df.columns.astype(str).str.strip()
                    if 'Route' in master_df.columns and 'Agency' in master_df.columns and 'DRCODE' in master_df.columns:
                        master_subset = master_df[['Route', 'Agency', 'DRCODE']].dropna(subset=['DRCODE'])
                        for _, r_item in master_subset.iterrows():
                            rt_key = str(r_item['Route']).replace('.0', '').strip()
                            ag_key = str(r_item['Agency']).replace('.0', '').strip()
                            mapping_dict[(rt_key, ag_key)] = str(r_item['DRCODE']).strip()

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
                    df_input_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)

                    fg_row, fg_col = -1, -1
                    for r in range(df_input_raw.shape[0]):
                        for c in range(df_input_raw.shape[1]):
                            val = str(df_input_raw.iloc[r, c]).strip().upper()
                            if "FG" in val:
                                fg_row, fg_col = r, c
                                break
                        if fg_row != -1:
                            break

                    if fg_row != -1:
                        route_num = "22"
                        match_route = re.search(r'Route\s*\(?(\d+)\)?', short_filename, re.IGNORECASE)
                        if match_route:
                            route_num = match_route.group(1)
                        else:
                            ignore_list = ["RT", "DR", "RT DR", "ROUTE", "SALES PERSON", "CONTACT NO:", "MATERIAL CODE"]
                            for r in range(fg_row):
                                for c in range(min(fg_col, 30)):
                                    cell_val = str(df_input_raw.iloc[r, c]).strip()
                                    upper_val = cell_val.upper()
                                    if upper_val in ignore_list or any(upper_val.startswith(p) for p in ["PC", "MS", "M", "GM", "DP", "SKU", "FG"]):
                                        continue
                                    if cell_val != "" and len(cell_val) <= 3 and any(char.isdigit() for char in cell_val):
                                        route_num = cell_val
                                        break
                                if route_num != "22":
                                    break

                        agency_col = -1
                        for cSearch in range(fg_col - 1, -1, -1):
                            valid_count = 0
                            for rCheck in range(fg_row + 1, min(fg_row + 15, df_input_raw.shape[0])):
                                v = df_input_raw.iloc[rCheck, cSearch]
                                if pd.notna(v):
                                    s_val = str(v).replace('.0', '').strip()
                                    if s_val.isdigit() and 1 <= len(s_val) <= 5:
                                        valid_count += 1
                            if valid_count >= 3:
                                agency_col = cSearch
                                break

                        if agency_col == -1 and fg_col > 0:
                            agency_col = fg_col - 1

                        wb_mod = openpyxl.load_workbook(io.BytesIO(file_bytes))
                        ws_mod = wb_mod.active
                        excel_fg_row = fg_row + 1
                        excel_fg_col = fg_col + 1

                        existing_drcode_col = None
                        for col_idx in range(1, ws_mod.max_column + 1):
                            cell_val = str(ws_mod.cell(row=excel_fg_row, column=col_idx).value).strip().upper()
                            if cell_val == "DRCODE":
                                existing_drcode_col = col_idx
                                break

                        target_col_idx = excel_fg_col
                        if existing_drcode_col:
                            target_col_idx = existing_drcode_col
                        else:
                            ws_mod.insert_cols(excel_fg_col)
                            target_col_idx = excel_fg_col
                            ws_mod.cell(row=excel_fg_row, column=target_col_idx, value="DRCODE")

                            for row in ws_mod.iter_rows():
                                for cell in row:
                                    if cell.value and str(cell.value).startswith('='):
                                        old_formula = str(cell.value)
                                        def shift_cols_in_formula(match):
                                            col_letters = match.group(1)
                                            row_num = match.group(2)
                                            col_idx = openpyxl.utils.column_index_from_string(col_letters)
                                            if col_idx >= target_col_idx:
                                                return f"{openpyxl.utils.get_column_letter(col_idx + 1)}{row_num}"
                                            return match.group(0)
                                        cell.value = re.sub(r'([A-Z]+)(\d+)', shift_cols_in_formula, old_formula)

                        for row_idx in range(excel_fg_row + 1, ws_mod.max_row + 1):
                            raw_agency = ws_mod.cell(row=row_idx, column=agency_col + 1).value
                            agency_str = str(raw_agency).replace('.0', '').strip() if raw_agency is not None else ""
                            if agency_str and agency_str != "None":
                                lookup_key = (str(route_num), agency_str)
                                assigned_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_str}")
                                ws_mod.cell(row=row_idx, column=target_col_idx, value=assigned_dr)

                        mod_buf = io.BytesIO()
                        wb_mod.save(mod_buf)
                        file_bytes = mod_buf.getvalue()

                    df_input = pd.read_excel(io.BytesIO(file_bytes), header=None)
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

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))
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
                            total_skipped_rows += 1
                            continue

                        agency_val = int(agency_str)
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

                        lookup_key = (str(route_num), str(agency_val))
                        has_dr_code = lookup_key in mapping_dict
                        clean_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_val}")

                        if not has_dr_code:
                            unmapped_records_to_store.append((short_filename, str(route_num), str(agency_val), clean_dr, batch_ts))
                            st.session_state.unmapped_current_batch.append({"File Name": short_filename, "Route": str(route_num), "Agency": agency_val, "Status": "Assigned NEW_CUST via Master lookup miss"})

                        if has_dr_code:
                            agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                            current_seq = agency_counts_valid[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_valid, valid_row, valid_order_num, clean_dr, "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            current_seq = agency_counts_missing[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-NEW-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_missing, missing_row, missing_order_num, clean_dr, "Missing DR"

                        item_id = 10
                        for c, fg_code, qty_val in valid_row_quantities:
                            current_fg = str(fg_code).strip() if str(fg_code).strip().upper().startswith("FG") else DEFAULTS["fg_code"]
                            total_input_qty += qty_val
                            total_gen_qty += qty_val

                            file_comparison_rows.append({"File Name": short_filename, "Status": file_category, "Agency": agency_val, "DR Code": dr_to_use, "FG Code": current_fg, "Input Qty": qty_val, "Generated Qty": qty_val})

                            target_ws.cell(row=current_r, column=2, value=order_num)
                            target_ws.cell(row=current_r, column=3, value="OR")
                            target_ws.cell(row=current_r, column=4, value="SO20")
                            target_ws.cell(row=current_r, column=5, value=10)
                            target_ws.cell(row=current_r, column=6, value=20)
                            target_ws.cell(row=current_r, column=7, value=dr_to_use)
                            target_ws.cell(row=current_r, column=8, value=dr_to_use)
                            target_ws.cell(row=current_r, column=9, value=ref_number)
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
                        out_fname = f"{safe_route_num}_{today_date}_{timestamp}_Valid.xlsx"
                        st.session_state.processed_files.append({"name": short_filename + " (Valid DR)", "data": buf_valid.getvalue(), "filename": out_fname, "orders": valid_items_created})
                        output_files_to_store.append((out_fname, "Valid DR", buf_valid.getvalue(), batch_ts))

                    if missing_items_created > 0:
                        buf_missing = io.BytesIO()
                        wb_missing.save(buf_missing)
                        out_fname_miss = f"{safe_route_num}_{today_date}_{timestamp}_Missing_DR.xlsx"
                        st.session_state.processed_files.append({"name": short_filename + " (Missing DR / New)", "data": buf_missing.getvalue(), "filename": out_fname_miss, "orders": missing_items_created})
                        output_files_to_store.append((out_fname_miss, "Missing DR", buf_missing.getvalue(), batch_ts))

                    if file_comparison_rows:
                        df_comp = pd.DataFrame(file_comparison_rows)
                        st.session_state.comparison_summary.append(df_comp.pivot_table(index=["File Name", "Status", "Agency", "DR Code", "FG Code"], values=["Input Qty", "Generated Qty"], aggfunc="sum").reset_index())

                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.executemany("INSERT OR IGNORE INTO unmapped_missing_dr_ledger (file_name, route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?, ?)", unmapped_records_to_store)
                for fname, ftype, fdata, fdate in output_files_to_store:
                    cursor.execute("INSERT OR REPLACE INTO output_files_ledger (file_name, file_type, file_data, created_at) VALUES (?, ?, ?, ?)", (fname, ftype, fdata, fdate))
                conn.commit()
                conn.close()

                st.session_state.kpi_data = {"input_qty": total_input_qty, "gen_qty": total_gen_qty, "valid_count": total_valid_orders, "missing_count": total_missing_orders, "skipped_count": total_skipped_rows}
                st.success("✅ Batch processing completed successfully using Master File mapping!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

# ==============================================================================
# SECTION 9: KPI METRIC CARDS & MULTI-CHANNEL DISPATCH HUB
# ==============================================================================
if st.session_state.processed_files:
    st.markdown("---")
    st.markdown("### 📈 Batch Performance & KPI Summary")
    kpi = st.session_state.kpi_data
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Input Qty", f"{kpi['input_qty']:,.0f}")
    col2.metric("Generated Qty", f"{kpi['gen_qty']:,.0f}")
    col3.metric("Valid Orders", kpi['valid_count'])
    col4.metric("Missing DR Orders", kpi['missing_count'])
    col5.metric("Skipped Rows", kpi['skipped_count'])

    st.markdown("---")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])

    col_zip, col_email, col_wa = st.columns(3)
    with col_zip:
        st.download_button("📦 Download ZIP of All Outputs", data=zip_buffer.getvalue(), file_name=f"Batch_Orders_{today_date}.zip", mime="application/zip")
    with col_email:
        if st.button("📧 Email Dispatch Reports"):
            st.success("Email feature ready (configure credentials in sidebar/code).")
    with col_wa:
        if whatsapp_num:
            wa_link = f"https://wa.me/{whatsapp_num}?text={urllib.parse.quote(f'Batch Orders Processed! Total Qty: {kpi[\"input_qty\"]}')}"
            st.markdown(f'<a href="{wa_link}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp Notification</button></a>', unsafe_allow_html=True)

# ==============================================================================
# SECTION 10: DATABASES & LEDGERS MANAGEMENT PANEL
# ==============================================================================
st.markdown("---")
with st.expander("🗄️ View & Manage Databases (Unmapped Ledger & Archived Outputs)"):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_unmapped = pd.read_sql("SELECT * FROM unmapped_missing_dr_ledger ORDER BY id DESC", conn)
        df_outputs = pd.read_sql("SELECT * FROM output_files_ledger ORDER BY id DESC", conn)
        conn.close()

        tab1, tab2 = st.tabs(["🚨 Unmapped Missing DR", "📦 Archived Outputs"])
        with tab1:
            if not df_unmapped.empty:
                st.dataframe(df_unmapped, use_container_width=True)
            else:
                st.info("No unmapped records found.")
        with tab2:
            if not df_outputs.empty:
                st.dataframe(df_outputs[['id', 'file_name', 'file_type', 'created_at']], use_container_width=True)
            else:
                st.info("No archived outputs found.")
    except Exception as e:
        st.error(f"Error loading database: {e}")
