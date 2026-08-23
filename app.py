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

# Page Configuration & Styling (Stretch & Alignment Fixed)
st.set_page_config(
    page_title="Sales Order Automation Hub", 
    page_icon="🚀", 
    layout="wide"
)

st.markdown("""
    <style>
        #GithubIcon { visibility: hidden !important; display: none !important; }
        .stAppHeader { background-color: transparent !important; }
        header[data-testid="stHeader"] { display: none !important; }
        
        /* Uniform button sizing and layout stabilization to prevent vertical stretching */
        .stButton>button {
            width: 100%;
            height: 50px;
            background-color: #10b981 !important;
            color: #ffffff !important;
            font-size: 15px;
            font-weight: 700;
            border-radius: 8px;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .stButton>button:hover {
            background-color: #059669 !important;
        }
    </style>
""", unsafe_allow_html=True)

# IST Timezone Helper via pytz
IST = pytz.timezone('Asia/Kolkata')

def get_ist_now():
    return datetime.datetime.now(IST)

# SQLite Database Initialization for Historical Trends (IST Timestamp)
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
    conn.commit()
    conn.close()

init_db()

# --- Session State Defaults for Reset/Clear/Restore ---
DEFAULTS = {
    "fg_code": "FG500014",
    "col_map": "36:FG500014AJ\n37:FG500014AK",
    "agency_override": "101:36:FG500014N01\n101:37:FG500014N02",
    "route": "22",
    "email_user": st.secrets.get("email", {}).get("sender_email", ""),
    "email_pass": st.secrets.get("email", {}).get("app_password", ""),
    "recipient": st.secrets.get("email", {}).get("recipient_email", ""),
    "whatsapp": ""
}

for key, val in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = val

# Sidebar Settings & Dynamic Mapping with Clear/Restore
st.sidebar.title("⚙️ System Settings")

if st.sidebar.button("🔄 Reset All to Defaults"):
    for k, v in DEFAULTS.items():
        st.session_state[k] = v
    st.rerun()

st.sidebar.markdown("---")

# 1. Default FG Code
st.sidebar.subheader("Default Fallback FG Code")
st.session_state.fg_code = st.sidebar.text_input("FG Code Input", value=st.session_state.fg_code, label_visibility="collapsed")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear FG"): st.session_state.fg_code = ""; st.rerun()
if c2.button("Restore FG"): st.session_state.fg_code = DEFAULTS["fg_code"]; st.rerun()

# 2. Column Mapping
st.sidebar.subheader("Direct Column Index Mapping")
st.session_state.col_map = st.sidebar.text_area("Col Map Input", value=st.session_state.col_map, label_visibility="collapsed", help="ColIndex:Code")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Map"): st.session_state.col_map = ""; st.rerun()
if c2.button("Restore Map"): st.session_state.col_map = DEFAULTS["col_map"]; st.rerun()

st.sidebar.markdown("---")

# 3. Agency Override
st.sidebar.subheader("Agency & Column-wise FG Override")
st.session_state.agency_override = st.sidebar.text_area("Agency Override Input", value=st.session_state.agency_override, label_visibility="collapsed", help="Agency:ColIndex:CustomFG")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Override"): st.session_state.agency_override = ""; st.rerun()
if c2.button("Restore Override"): st.session_state.agency_override = DEFAULTS["agency_override"]; st.rerun()

# 4. Route Fallback
st.sidebar.subheader("Default Route Fallback")
st.session_state.route = st.sidebar.text_input("Route Input", value=st.session_state.route, label_visibility="collapsed")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Route"): st.session_state.route = ""; st.rerun()
if c2.button("Restore Route"): st.session_state.route = DEFAULTS["route"]; st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📧 Email Dispatch Settings")

# 5. Sender Email
st.sidebar.text("Sender Email ID")
st.session_state.email_user = st.sidebar.text_input("Sender Input", value=st.session_state.email_user, label_visibility="collapsed")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Email"): st.session_state.email_user = ""; st.rerun()
if c2.button("Restore Email"): st.session_state.email_user = DEFAULTS["email_user"]; st.rerun()

# 6. Email Password
st.sidebar.text("Email App Password")
st.session_state.email_pass = st.sidebar.text_input("Pass Input", type="password", value=st.session_state.email_pass, label_visibility="collapsed")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Pass"): st.session_state.email_pass = ""; st.rerun()
if c2.button("Restore Pass"): st.session_state.email_pass = DEFAULTS["email_pass"]; st.rerun()

# 7. Recipient Email
st.sidebar.text("Recipient Email")
st.session_state.recipient = st.sidebar.text_input("Recipient Input", value=st.session_state.recipient, label_visibility="collapsed")
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear Recipient"): st.session_state.recipient = ""; st.rerun()
if c2.button("Restore Recipient"): st.session_state.recipient = DEFAULTS["recipient"]; st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📱 WhatsApp Notification")

# 8. WhatsApp Number
st.session_state.whatsapp = st.sidebar.text_input("WhatsApp Number (e.g., 919876543210)", value=st.session_state.whatsapp)
c1, c2 = st.sidebar.columns(2)
if c1.button("Clear WA"): st.session_state.whatsapp = ""; st.rerun()
if c2.button("Restore WA"): st.session_state.whatsapp = DEFAULTS["whatsapp"]; st.rerun()

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

st.title("🚀 Enterprise Sales Order Automation Hub (Pro Edition)")
st.markdown("Upload multiple **Inbound Demand Files** to process orders, track advanced KPIs, run smart anomaly detection, and export audit reports.")
st.markdown("---")

# Session State Initialization
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'comparison_summary' not in st.session_state:
    st.session_state.comparison_summary = []
if 'skipped_rows_log' not in st.session_state:
    st.session_state.skipped_rows_log = []
if 'kpi_data' not in st.session_state:
    st.session_state.kpi_data = {"input_qty": 0, "gen_qty": 0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

# --- Pre-flight File Health Check ---
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

if st.button("🚀 Process Batch Orders & Audit Logs", type="primary"):
    if uploaded_inputs:
        st.session_state.processed_files = []
        st.session_state.comparison_summary = []
        st.session_state.skipped_rows_log = []
        
        total_input_qty = 0
        total_gen_qty = 0
        total_valid_orders = 0
        total_missing_orders = 0
        total_skipped_rows = 0
        
        with st.spinner("⚡ Reading files, processing orders, and checking for exceptions... Please wait."):
            try:
                try:
                    with open("Output.xlsx", "rb") as f:
                        template_bytes = f.read()
                except FileNotFoundError:
                    st.error("❌ 'Output.xlsx' template file repository mein nahi mili. Kripya template file ko GitHub repo ke main folder mein upload karein.")
                    st.stop()
                
                ist_now = get_ist_now()
                today_date = ist_now.strftime("%Y-%m-%d")
                timestamp = ist_now.strftime("%H%M%S")

                for uploaded_file in uploaded_inputs:
                    short_filename = uploaded_file.name
                    if short_filename.lower() == "output.xlsx":
                        continue

                    file_bytes = uploaded_file.getvalue()
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

                    # 2. Strict Total/Sum Column Detection
                    total_col = df_input.shape[1]
                    for cSearch in range(fg_col, df_input.shape[1]):
                        is_total = False
                        for scan_r in range(max(0, fg_row - 10), min(fg_row + 3, df_input.shape[0])):
                            cell_val = str(df_input.iloc[scan_r, cSearch]).strip().upper()
                            if any(kw in cell_val for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                                is_total = True
                                break
                            if scan_r >= fg_row + 1 and ("SUM" in cell_val or "=" in cell_val):
                                is_total = True
                                break
                        if is_total:
                            total_col = cSearch
                            break

                    # 3. Route Number Finding Logic
                    route_num = default_fallback_route if default_fallback_route != "" else "22"
                    ignore_list = ["RT", "DR", "RT DR", "ROUTE", "SALES PERSON", "CONTACT NO:", "MATERIAL CODE"]
                    
                    for r in range(fg_row):
                        for c in range(min(total_col, 30)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            upper_val = cell_val.upper()
                            if upper_val in ignore_list:
                                continue
                            is_product_code = any(upper_val.startswith(p) for p in ["PC", "MS", "M", "GM", "DP", "SKU", "FG"])
                            if is_product_code:
                                continue
                            if cell_val != "" and 1 <= len(cell_val) <= 3:
                                if any(char.isdigit() for char in cell_val):
                                    route_num = cell_val
                                    break
                        if route_num != (default_fallback_route if default_fallback_route != "" else "22"):
                            break

                    if default_fallback_route != "" and default_fallback_route != "22":
                        route_num = default_fallback_route

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 4. Smart Agency Detection
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

                    # 4.1 Strict DR Code Column Detection
                    dr_code_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        sample_val = str(df_input.iloc[fg_row + 1, cSearch] if fg_row + 1 < df_input.shape[0] else "").strip().upper()
                        if re.match(r'^DR\d+', sample_val):
                            dr_code_col = cSearch
                            break
                        matched_count = 0
                        for offset in range(1, min(4, df_input.shape[0] - fg_row)):
                            v = str(df_input.iloc[fg_row + offset, cSearch]).strip().upper()
                            if re.match(r'^DR\d+', v):
                                matched_count += 1
                        if matched_count > 0:
                            dr_code_col = cSearch
                            break

                    # 5. Pure Valid FG Columns Collection
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

                        # Quantities Check
                        row_has_items = False
                        valid_row_quantities = []
                        for c, fg_code in valid_cols:
                            if c >= total_col:
                                continue
                            sku_qty = df_input.iloc[r, c]
                            if pd.notna(sku_qty) and str(sku_qty).strip() != "":
                                try:
                                    qty_val = float(sku_qty)
                                    if qty_val > 0:
                                        row_has_items = True
                                        valid_row_quantities.append((c, fg_code, qty_val))
                                except ValueError:
                                    pass

                        if not row_has_items:
                            st.session_state.skipped_rows_log.append({
                                "File Name": short_filename,
                                "Row Index": r + 1,
                                "Agency Value": agency_val,
                                "Reason": "Skipped: Zero or Blank Quantities across all SKUs"
                            })
                            total_skipped_rows += 1
                            continue

                        if has_dr_code:
                            agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                            current_seq = agency_counts_valid[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_valid, valid_row, valid_order_num, clean_dr, "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            current_seq = agency_counts_missing[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-NEW-{current_seq}"
                            target_ws, current_r, order_num, dr_to_use, file_category = ws_missing, missing_row, missing_order_num, f"NEW_CUST_{agency_val}", "Missing DR"

                        item_id = 10
                        for c, fg_code, qty_val in valid_row_quantities:
                            cleaned_fg = str(fg_code).strip()
                            upper_fg = cleaned_fg.upper()
                            
                            if (agency_val, c) in agency_col_override_map:
                                current_fg = agency_col_override_map[(agency_val, c)]
                            elif upper_fg.startswith("FG"):
                                current_fg = cleaned_fg
                            elif upper_fg in ["", "NAN", "NONE"]:
                                current_fg = direct_col_mapping.get(c, default_fg_code)
                            else:
                                current_fg = direct_col_mapping.get(c, default_fg_code)
                            
                            total_input_qty += qty_val
                            total_gen_qty += qty_val
                            
                            file_comparison_rows.append({
                                "File Name": short_filename,
                                "Status": file_category,
                                "Agency": agency_val,
                                "DR Code": dr_to_use,
                                "FG Code": current_fg,
                                "Input Qty": qty_val,
                                "Generated Qty": qty_val
                            })
                            
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
                        buf_valid.seek(0)
                        st.session_state.processed_files.append({
                            "name": short_filename + " (Valid DR)",
                            "data": buf_valid.getvalue(),
                            "filename": safe_route_num + "_" + today_date + "_" + timestamp + "_Valid.xlsx",
                            "orders": valid_items_created
                        })

                    if missing_items_created > 0:
                        buf_missing = io.BytesIO()
                        wb_missing.save(buf_missing)
                        buf_missing.seek(0)
                        st.session_state.processed_files.append({
                            "name": short_filename + " (Missing DR / New)",
                            "data": buf_missing.getvalue(),
                            "filename": safe_route_num + "_" + today_date + "_" + timestamp + "_Missing_DR.xlsx",
                            "orders": missing_items_created
                        })

                    if file_comparison_rows:
                        df_comp = pd.DataFrame(file_comparison_rows)
                        df_pivot = df_comp.pivot_table(
                            index=["File Name", "Status", "Agency", "DR Code", "FG Code"],
                            values=["Input Qty", "Generated Qty"],
                            aggfunc="sum"
                        ).reset_index()
                        df_pivot["Difference"] = df_pivot["Input Qty"] - df_pivot["Generated Qty"]
                        st.session_state.comparison_summary.append(df_pivot)

                # Store KPI metrics & Save to SQLite Database (IST Timestamp)
                st.session_state.kpi_data = {
                    "input_qty": total_input_qty,
                    "gen_qty": total_gen_qty,
                    "valid_count": total_valid_orders,
                    "missing_count": total_missing_orders,
                    "skipped_count": total_skipped_rows
                }
                
                conn = sqlite3.connect("sales_history.db")
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO history_logs (timestamp, files_count, total_qty, status) VALUES (?, ?, ?, ?)",
                    (get_ist_now().strftime("%Y-%m-%d %H:%M:%S"), len(uploaded_inputs), total_input_qty, "Success")
                )
                conn.commit()
                conn.close()

                st.success("✅ Batch Processing & Advanced Audit Complete!")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.warning("⚠️ Kripya pehle demand files upload karein!")

# Display KPI Summary Cards & Advanced Visual Analytics
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

    if kpi['skipped_count'] > 5:
        st.warning(f"⚠️ **Smart Audit Alert:** {kpi['skipped_count']} rows skipped check exception logs.")

    # --- ADVANCED TABBED VISUAL ANALYTICS ---
    if st.session_state.comparison_summary:
        st.markdown("---")
        st.markdown("### 📊 Advanced Visual Analytics Dashboard")
        combined_df_chart = pd.concat(st.session_state.comparison_summary, ignore_index=True)
        
        tab1, tab2 = st.tabs(["📊 Agency-wise Breakdown", "📦 SKU-wise Share"])
        with tab1:
            st.bar_chart(combined_df_chart.groupby("Agency")["Generated Qty"].sum())
        with tab2:
            st.bar_chart(combined_df_chart.groupby("FG Code")["Generated Qty"].sum())

    st.markdown("---")
    st.markdown("### 📥 Bulk Download & Notifications")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])
    
    col_zip, col_pdf, col_summary, col_json, col_print, col_email, col_wa = st.columns(7)
    
    with col_zip:
        st.download_button(
            label="📦 ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"Batch_Orders_{get_ist_now().strftime('%Y-%m-%d')}.zip",
            mime="application/zip",
            key="zip_download"
        )
        
    with col_pdf:
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(190, 10, "Enterprise Sales Order Summary Invoice", ln=True, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(190, 6, f"Generated On (IST): {get_ist_now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
            pdf.ln(10)
            
            pdf.set_font("Arial", "B", 11)
            pdf.cell(100, 8, "Metric Description", border=1)
            pdf.cell(90, 8, "Value", border=1, ln=True)
            
            pdf.set_font("Arial", "", 11)
            metrics_list = [
                ("Total Input Quantity", f"{kpi['input_qty']:,.0f}"),
                ("Total Generated Quantity", f"{kpi['gen_qty']:,.0f}"),
                ("Valid DR Orders", str(kpi['valid_count'])),
                ("Missing DR Orders", str(kpi['missing_count'])),
                ("Skipped Rows Logged", str(kpi['skipped_count'])),
                ("Success Rate", f"{success_rate:.1f}%")
            ]
            for m_desc, m_val in metrics_list:
                pdf.cell(100, 8, m_desc, border=1)
                pdf.cell(90, 8, m_val, border=1, ln=True)
                
            pdf_bytes = bytes(pdf.output())
            st.download_button(
                label="📄 PDF",
                data=pdf_bytes,
                file_name=f"Sales_Invoice_{get_ist_now().strftime('%Y-%m-%d')}.pdf",
                mime="application/pdf",
                key="pdf_invoice_download"
            )
        except Exception as e:
            st.error(f"PDF Error: {str(e)}")

    with col_summary:
        summary_txt = f"""=== ENTERPRISE SALES ORDER SUMMARY (IST) ===
Date/Time: {get_ist_now().strftime('%Y-%m-%d %H:%M:%S')}
----------------------------------------
Total Input Quantity : {kpi['input_qty']:,.0f}
Total Generated Qty  : {kpi['gen_qty']:,.0f}
Valid DR Orders      : {kpi['valid_count']}
Missing DR Orders    : {kpi['missing_count']}
Skipped Rows Logged  : {kpi['skipped_count']}
Success Rate         : {success_rate:.1f}%
----------------------------------------
Generated Files Count: {len(st.session_state.processed_files)}
Status: Successfully Processed & Audited
========================================"""
        st.download_button(
            label="📄 TXT",
            data=summary_txt.encode('utf-8'),
            file_name=f"Summary_Report_{get_ist_now().strftime('%Y-%m-%d')}.txt",
            mime="text/plain",
            key="summary_txt_download"
        )
        
    with col_json:
        json_data = json.dumps({
            "timestamp": get_ist_now().strftime('%Y-%m-%d %H:%M:%S'),
            "metrics": kpi,
            "success_rate": f"{success_rate:.1f}%"
        }, indent=4)
        st.download_button(
            label="💾 JSON",
            data=json_data.encode('utf-8'),
            file_name=f"Audit_Backup_{get_ist_now().strftime('%Y-%m-%d')}.json",
            mime="application/json",
            key="json_backup_download"
        )
        
    with col_print:
        print_html = """
        <div style="width:100%; margin:0; padding:0;">
            <button onclick="parent.window.print()" style="width:100%; height:50px; background:#3b82f6; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer; font-family:sans-serif; display:flex; align-items:center; justify-content:center;">
                🖨️ Print
            </button>
        </div>
        """
        components.html(print_html, height=60)
        
    with col_email:
        if st.button("📧 Email"):
            if email_user and email_pass and recipient_email:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = f"🚀 Sales Orders Batch Execution Report (IST) - {get_ist_now().strftime('%Y-%m-%d')}"
                    msg['From'] = email_user
                    msg['To'] = recipient_email
                    
                    # --- Enhanced Modern Rich HTML Email Table ---
                    html_content = f"""
                    <html>
                      <body style="font-family: Arial, sans-serif; color: #333; background-color: #f9fafb; padding: 20px;">
                        <div style="max-width: 600px; background: #ffffff; padding: 25px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                          <h2 style="color: #10b981; border-bottom: 2px solid #e5e7eb; padding-bottom: 10px;">📊 Sales Order Batch Automation Hub</h2>
                          <p>Hello Team,</p>
                          <p>The daily inbound demand batch has been processed successfully on <b>{get_ist_now().strftime('%Y-%m-%d %H:%M:%S')} IST</b>.</p>
                          <table style="border-collapse: collapse; width: 100%; margin-top: 15px; border-radius: 6px; overflow: hidden;">
                            <tr style="background-color: #10b981; color: white;">
                              <th style="padding: 10px; text-align: left;">Metric</th>
                              <th style="padding: 10px; text-align: left;">Value</th>
                            </tr>
                            <tr style="background-color: #f3f4f6;">
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Total Input Qty</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;"><b>{kpi['input_qty']:,.0f}</b></td>
                            </tr>
                            <tr>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Valid Orders</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">{kpi['valid_count']}</td>
                            </tr>
                            <tr style="background-color: #f3f4f6;">
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Missing DR Orders</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">{kpi['missing_count']}</td>
                            </tr>
                            <tr>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">Success Rate</td>
                              <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;"><b>{success_rate:.1f}%</b></td>
                            </tr>
                            <tr style="background-color: #f3f4f6;">
                              <td style="padding: 10px;">Skipped Rows</td>
                              <td style="padding: 10px;">{kpi['skipped_count']}</td>
                            </tr>
                          </table>
                          <p style="margin-top: 25px; color: #666; font-size: 12px; border-top: 1px solid #e5e7eb; paddingTop: 10px;">Automated via Sales Order Hub (IST)</p>
                        </div>
                      </body>
                    </html>
                    """
                    msg.set_content("Please enable HTML to view this report.")
                    msg.add_alternative(html_content, subtype='html')
                    
                    for item in st.session_state.processed_files:
                        msg.add_attachment(item['data'], maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=item['filename'])
                    
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email_user, email_pass)
                        smtp.send_message(msg)
                    st.success("✅ Email dispatched!")
                except Exception as e:
                    st.error(f"❌ Email failed: {str(e)}")
            else:
                st.warning("⚠️ Enter email credentials!")

    with col_wa:
        if whatsapp_num:
            wa_text = f"Sales Order Batch Ready! Total Qty: {kpi['input_qty']}, Success Rate: {success_rate:.1f}%."
            wa_link = f"https://wa.me/{whatsapp_num}?text={urllib.parse.quote(wa_text)}"
            st.markdown(f'<a href="{wa_link}" target="_blank" style="text-decoration:none;"><button style="width:100%; height:50px; background:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer; display:flex; align-items:center; justify-content:center;">📱 WhatsApp</button></a>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("##### Individual File Downloads:")
    for i, item in enumerate(st.session_state.processed_files):
        if st.download_button(
            label=f"📥 Download {item['name']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_file_{i}_{item['filename']}"
        ):
            st.toast(f"🎉 '{item['filename']}' downloaded!", icon="📥")

# Tables and Logs with Search & Filter Feature
if st.session_state.comparison_summary:
    st.markdown("---")
    st.markdown("### 📋 Agency-wise Material & Input Comparison")
    
    search_query = st.text_input("🔍 Search Table (Filter by Agency, DR Code, or FG Code)", "", key="table_search")
    
    combined_df = pd.concat(st.session_state.comparison_summary, ignore_index=True)
    summary_table = combined_df.groupby(["Agency", "DR Code", "FG Code"], as_index=False).agg({"Input Qty": "sum", "Generated Qty": "sum"})
    summary_table["Difference"] = summary_table["Input Qty"] - summary_table["Generated Qty"]
    
    if search_query:
        q = search_query.lower()
        summary_table = summary_table[
            summary_table['Agency'].astype(str).str.lower().str.contains(q) |
            summary_table['DR Code'].astype(str).str.lower().str.contains(q) |
            summary_table['FG Code'].astype(str).str.lower().str.contains(q)
        ]
        
    st.dataframe(summary_table, use_container_width=True)

if st.session_state.skipped_rows_log:
    st.markdown("---")
    st.markdown("### ⚠️ Skipped / Invalid Rows Exception Log")
    df_skipped = pd.DataFrame(st.session_state.skipped_rows_log)
    
    skip_search = st.text_input("🔍 Search Skipped Log (Filter by File Name or Agency)", "", key="skip_search")
    if skip_search:
        sq = skip_search.lower()
        df_skipped = df_skipped[
            df_skipped['File Name'].astype(str).str.lower().str.contains(sq) |
            df_skipped['Agency Value'].astype(str).str.lower().str.contains(sq)
        ]
    st.dataframe(df_skipped, use_container_width=True)

if st.session_state.comparison_summary:
    st.markdown("---")
    st.markdown("### 📊 Audit Reconciliation & Comparison Pivot")
    combined_pivot = pd.concat(st.session_state.comparison_summary, ignore_index=True)
    st.dataframe(combined_pivot, use_container_width=True)
    audit_csv = combined_pivot.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Audit Reconciliation Report (CSV)",
        data=audit_csv,
        file_name=f"Audit_Reconciliation_Report_{get_ist_now().strftime('%Y-%m-%d')}.csv",
        mime="text/csv",
        key="audit_csv_download"
    )

# --- Historical Trend Analysis View ---
st.markdown("---")
with st.expander("🕒 View Historical Trend Analysis (SQLite Database - IST)"):
    try:
        conn = sqlite3.connect("sales_history.db")
        df_history = pd.read_sql("SELECT * FROM history_logs ORDER BY id DESC", conn)
        conn.close()
        if not df_history.empty:
            st.dataframe(df_history, use_container_width=True)
            st.markdown("##### Day-over-Day Total Quantity Trend")
            st.line_chart(df_history.set_index("timestamp")["total_qty"])
        else:
            st.info("No historical logs available yet.")
    except Exception as e:
        st.error(f"Error loading history: {str(e)}")