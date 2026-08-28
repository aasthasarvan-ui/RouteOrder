import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import datetime
import pytz
import io
import sqlite3
import smtplib
from email.message import EmailMessage

# ==============================================================================
# PAGE CONFIGURATION & TIMEZONE
# ==============================================================================
try:
    st.set_page_config(
        page_title="Enterprise Vehicle Tonnage Hub",
        page_icon="🚚",
        layout="wide",
        initial_sidebar_state="expanded"
    )
except:
    pass

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

# ==============================================================================
# DATABASE INITIALIZATION (PERMANENT BLOB STORAGE)
# ==============================================================================
def init_db():
    try:
        conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS saved_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                upload_date TEXT,
                file_name TEXT,
                file_blob BLOB
            )
        """)
        conn.commit()
        conn.close()
    except:
        pass

init_db()

# ==============================================================================
# SESSION STATE MANAGEMENT
# ==============================================================================
if "active_df" not in st.session_state:
    st.session_state.active_df = None
if "selected_file_id" not in st.session_state:
    st.session_state.selected_file_id = None
if "calc_theme_choice" not in st.session_state:
    st.session_state.calc_theme_choice = "⚡ Cyber Neon Glass"
if "ng_reset_token" not in st.session_state:
    st.session_state.ng_reset_token = 0

# ==============================================================================
# SMART DATAFRAME CLEANING & HEADER DETECTION
# ==============================================================================
def load_and_clean_dataframe(file_bytes, file_name):
    if file_name.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        excel_obj = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_target = 'Sheet1' if 'Sheet1' in excel_obj.sheet_names else excel_obj.sheet_names[0]
        df = pd.read_excel(excel_obj, sheet_name=sheet_target)
    
    # Check if header is misplaced (e.g. Unnamed columns)
    if any(str(c).startswith("Unnamed") for c in df.columns):
        for idx, row in df.head(10).iterrows():
            row_str = str(row.values).lower()
            if "vehicle" in row_str or "quantity" in row_str or "billing type" in row_str:
                df.columns = df.iloc[idx]
                df = df.iloc[idx+1:].reset_index(drop=True)
                break
    
    # Drop completely empty rows/columns
    df = df.dropna(how='all').reset_index(drop=True)
    return df

# ==============================================================================
# UI STYLING & HIGH-CONTRAST WHITE TEXT LIVE CLOCK HEADER
# ==============================================================================
st.markdown("""
    <style>
        .stApp {
            background-color: #0f172a;
            color: #f8fafc;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .main-hero {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            padding: 24px;
            border-radius: 16px;
            border: 1px solid #334155;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4);
            margin-bottom: 20px;
        }
        .live-clock-box {
            background: #1e293b;
            color: #ffffff !important;
            padding: 14px 22px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: 700;
            border: 2px solid #38bdf8;
            display: inline-block;
            text-align: center;
            box-shadow: 0 4px 15px rgba(56, 189, 248, 0.3);
            width: 100%;
        }
    </style>
""", unsafe_allow_html=True)

clock_html = """
    <div style="width: 100%;">
        <div class="live-clock-box" id="liveClock" style="color: #ffffff !important;">🕒 Initializing Clock...</div>
    </div>
    <script>
        function updateClock() {
            const now = new Date();
            const options = { timeZone: 'Asia/Kolkata', hour12: true, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
            const clockEl = document.getElementById('liveClock');
            clockEl.innerHTML = '🕒 ' + now.toLocaleString('en-IN', options) + ' IST';
            clockEl.style.color = '#ffffff';
        }
        setInterval(updateClock, 1000);
        updateClock();
    </script>
"""

col_h1, col_h2 = st.columns([2, 1])
with col_h1:
    st.markdown("""
        <div class="main-hero" style="margin-bottom:0px; padding:18px;">
            <h2 style="color: #f8fafc; margin: 0;">🚚 Enterprise Vehicle Tonnage Hub</h2>
            <p style="color: #94a3b8; margin: 6px 0 0 0; font-size: 13px;">
                Permanent Vault Storage, Vehicle Tonnage (F2 & BAG), Manual Calculator & Excel/Email Export.
            </p>
        </div>
    """, unsafe_allow_html=True)
with col_h2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.components.v1.html(clock_html, height=75)

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# LEFT SIDEBAR: DYNAMIC VAULT UPLOADER & SAVED FILES SELECTOR
# ==============================================================================
with st.sidebar:
    st.markdown("### 📂 Upload New Billing Export")
    uploaded_file = st.file_uploader("Upload File (.xlsx / .csv)", type=["xlsx", "csv"], key="sidebar_uploader")

    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.getvalue()
            temp_df = load_and_clean_dataframe(file_bytes, uploaded_file.name)

            if st.button("💾 Save to Permanent Vault", type="primary"):
                conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
                upload_timestamp = get_ist_now().strftime('%Y-%m-%d %H:%M:%S')
                cursor = conn.cursor()
                
                cursor.execute(
                    "INSERT INTO saved_files (upload_date, file_name, file_blob) VALUES (?, ?, ?)",
                    (upload_timestamp, uploaded_file.name, sqlite3.Binary(file_bytes))
                )
                conn.commit()
                conn.close()
                st.session_state.active_df = temp_df
                st.success(f"✅ '{uploaded_file.name}' saved successfully in Vault!")
                st.rerun()
        except Exception as err_file:
            st.error(f"❌ Upload error: {str(err_file)}")

    st.markdown("---")
    st.markdown("### 🗂️ Select File from Vault (No Re-upload needed)")
    
    try:
        conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
        saved_records = pd.read_sql("SELECT id, upload_date, file_name FROM saved_files ORDER BY id DESC", conn)
        conn.close()
    except:
        saved_records = pd.DataFrame()

    if not saved_records.empty:
        file_options = {f"[{row['id']}] {row['file_name']} ({row['upload_date']})": row['id'] for _, row in saved_records.iterrows()}
        selected_file_label = st.selectbox("Choose Saved File:", options=list(file_options.keys()))
        
        selected_id = file_options[selected_file_label]
        
        col_load, col_del = st.columns(2)
        with col_load:
            if st.button("📂 Load File", type="primary"):
                try:
                    conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
                    cursor = conn.cursor()
                    cursor.execute("SELECT file_blob, file_name FROM saved_files WHERE id = ?", (selected_id,))
                    row_data = cursor.fetchone()
                    conn.close()
                    if row_data:
                        blob_data, fname = row_data
                        df_from_db = load_and_clean_dataframe(blob_data, fname)
                        st.session_state.active_df = df_from_db
                        st.success(f"✅ Loaded '{fname}' successfully!")
                        st.rerun()
                except Exception as load_err:
                    st.error(f"Load error: {load_err}")
        
        with col_del:
            if st.button("🗑️ Delete File", type="secondary"):
                try:
                    conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM saved_files WHERE id = ?", (selected_id,))
                    conn.commit()
                    conn.close()
                    st.session_state.active_df = None
                    st.success("🗑️ File deleted from vault!")
                    st.rerun()
                except Exception as del_err:
                    st.error(f"Delete error: {del_err}")
    else:
        st.info("No saved files in vault. Upload a file above.")

# ==============================================================================
# AUTO-LOAD LATEST FILE FROM VAULT IF SESSION IS EMPTY
# ==============================================================================
if st.session_state.active_df is None and not saved_records.empty:
    try:
        latest_id = saved_records.iloc[0]['id']
        conn = sqlite3.connect("tonnage_master_hub.db", check_same_thread=False)
        cursor = conn.cursor()
        cursor.execute("SELECT file_blob, file_name FROM saved_files WHERE id = ?", (latest_id,))
        row_data = cursor.fetchone()
        conn.close()
        if row_data:
            blob_data, fname = row_data
            df_from_db = load_and_clean_dataframe(blob_data, fname)
            st.session_state.active_df = df_from_db
    except:
        pass

# ==============================================================================
# MAIN DASHBOARD & TONNAGE CALCULATORS
# ==============================================================================
if st.session_state.active_df is not None:
    st.markdown("### 📊 Billing & Tonnage Data Dashboard")

    working_df = st.session_state.active_df.copy()

    global_search = st.text_input("🔍 Global Keyword Search (Vehicle, Customer, Material, Document):", "", key="global_search_input")
    if str(global_search).strip() != "":
        term = str(global_search).strip().lower()
        mask = pd.Series(False, index=working_df.index)
        for c in working_df.columns:
            mask = mask | working_df[c].astype(str).str.lower().str.contains(term, na=False)
        working_df = working_df[mask]

    # ==========================================================================
    # VEHICLE TONNAGE CALCULATOR
    # ==========================================================================
    with st.expander("🚚 Vehicle Tonnage Calculator (Billing Type F2 & BAG Slabs)", expanded=True):
        st.markdown("Select Vehicle Number and Billing Date to instantly compute exact Precision vs Standard Tonnage based on your billing records.")
        
        all_cols_list = [str(c) for c in working_df.columns]
        
        default_veh_idx, default_qty_idx = 0, 0
        for idx, c in enumerate(all_cols_list):
            c_l = c.lower()
            if "vehicle" in c_l or "truck" in c_l or "veh" in c_l:
                default_veh_idx = idx
            if "quantity" in c_l or "qty" in c_l:
                default_qty_idx = idx

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            veh_col = st.selectbox("📌 Select Vehicle Column:", options=all_cols_list, index=default_veh_idx, key="sel_veh_col_map")
        with col_m2:
            qty_col = st.selectbox("📌 Select Quantity Column:", options=all_cols_list, index=default_qty_idx, key="sel_qty_col_map")

        date_col, btype_col, unit_col, desc_col = None, None, None, None
        for c in working_df.columns:
            c_low = str(c).lower()
            if not date_col and any(k in c_low for k in ["billing date", "date", "inv date", "posting"]):
                date_col = c
            if not btype_col and any(k in c_low for k in ["billing type", "btype", "type"]):
                btype_col = c
            if not unit_col and any(k in c_low for k in ["sale unit", "unit", "uom"]):
                unit_col = c
            if not desc_col and any(k in c_low for k in ["product description", "material description", "desc", "item description", "text"]):
                desc_col = c

        if veh_col and qty_col:
            calc_df = working_df.copy()
            if btype_col:
                calc_df = calc_df[calc_df[btype_col].astype(str).str.upper().str.contains("F2", na=False)]
            if unit_col:
                calc_df = calc_df[calc_df[unit_col].astype(str).str.upper().str.contains("BAG", na=False)]

            # Filter out rows where vehicle name looks like a summary total
            if not calc_df.empty:
                calc_df = calc_df[~calc_df[veh_col].astype(str).str.lower().str.contains("total", na=False)]

            unique_vehicles = sorted(calc_df[veh_col].dropna().astype(str).unique().tolist()) if not calc_df.empty else []
            
            if unique_vehicles:
                sel_vehicle = st.selectbox("🚛 Select Vehicle Number:", options=unique_vehicles, key="calc_veh_select")
                
                veh_subset = calc_df[calc_df[veh_col].astype(str) == sel_vehicle]
                unique_dates = sorted(veh_subset[date_col].dropna().astype(str).unique().tolist()) if date_col and not veh_subset.empty else ["All Dates"]
                
                sel_date = st.selectbox("📅 Select Billing Date:", options=unique_dates, key="calc_date_select")
                
                if date_col and sel_date != "All Dates":
                    final_veh_rows = veh_subset[veh_subset[date_col].astype(str) == sel_date]
                else:
                    final_veh_rows = veh_subset

                b50_total = 0.0
                b25_total = 0.0
                
                for _, r in final_veh_rows.iterrows():
                    val = r[qty_col]
                    try:
                        q = float(val) if pd.notna(val) else 0.0
                    except (ValueError, TypeError):
                        q = 0.0
                        
                    d_text = str(r[desc_col]).upper() if desc_col and pd.notna(r[desc_col]) else ""
                    
                    if "25" in d_text:
                        b25_total += q
                    else:
                        b50_total += q

                w1_opt1 = b50_total * 50.120 + b25_total * 25.120
                mt1 = w1_opt1 / 1000.0

                w1_opt2 = b50_total * 50.0 + b25_total * 25.0
                mt2 = w1_opt2 / 1000.0

                st.markdown(f"**Vehicle Summary for `{sel_vehicle}` (Total Bags: {int(b50_total + b25_total):,}):**")
                c_res1, c_res2 = st.columns(2)
                with c_res1:
                    st.metric("🔹 Precision Scale (Opt 1: 50.12 & 25.12)", f"{mt1:,.3f} MT", f"{w1_opt1:,.2f} Kgs")
                with c_res2:
                    st.metric("🔹 Standard Slabs (Opt 2: 50.0 & 25.0)", f"{mt2:,.3f} MT", f"{w1_opt2:,.2f} Kgs")
            else:
                st.info("ℹ️ No records found matching Billing Type 'F2' and Unit 'BAG'.")
        else:
            st.warning("⚠️ Please select valid Vehicle and Quantity columns above.")

    # ======================================================================
    # MANUAL BAG INPUT CALCULATOR
    # ======================================================================
    with st.expander("⚡ Next-Gen Enterprise Manual Tonnage Calculator", expanded=False):
        CALC_THEMES = {
            "⚡ Cyber Neon Glass": {
                "bg": "#0b0f19", "card_bg": "rgba(30, 41, 59, 0.75)", "border": "#38bdf8", 
                "accent": "#38bdf8", "text": "#f8fafc", "subtext": "#94a3b8", "radius": "16px"
            },
            "🌲 Emerald Corporate": {
                "bg": "#064e3b", "card_bg": "rgba(6, 78, 59, 0.85)", "border": "#34d399", 
                "accent": "#34d399", "text": "#ecfdf5", "subtext": "#a7f3d0", "radius": "10px"
            },
            "🍇 Executive Burgundy": {
                "bg": "#500724", "card_bg": "rgba(80, 7, 36, 0.85)", "border": "#f472b6", 
                "accent": "#f472b6", "text": "#fdf2f8", "subtext": "#fbcfe8", "radius": "20px"
            },
            "💼 Classic Navy Slate": {
                "bg": "#0f172a", "card_bg": "rgba(15, 23, 42, 0.9)", "border": "#64748b", 
                "accent": "#60a5fa", "text": "#ffffff", "subtext": "#cbd5e1", "radius": "12px"
            }
        }

        selected_theme_name = st.selectbox("🎨 Select Calculator UI Theme & Layout Style:", list(CALC_THEMES.keys()), key="calc_theme_dropdown")
        st.session_state.calc_theme_choice = selected_theme_name
        th = CALC_THEMES[selected_theme_name]

        st.markdown(f"""
            <style>
                .nextgen-card {{
                    background: {th['card_bg']};
                    backdrop-filter: blur(12px);
                    padding: 22px;
                    border-radius: {th['radius']};
                    border: 1px solid {th['border']};
                    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.2);
                    margin-bottom: 15px;
                }}
                .card-header {{
                    color: {th['accent']};
                    font-size: 17px;
                    font-weight: 700;
                    margin-bottom: 12px;
                }}
                .metric-value {{
                    font-size: 26px;
                    font-weight: 800;
                    color: {th['text']};
                    margin: 6px 0;
                }}
                .sub-metric {{
                    font-size: 13px;
                    color: {th['subtext']};
                }}
            </style>
        """, unsafe_allow_html=True)

        k_f1 = f"ng_input_50kg_{st.session_state.ng_reset_token}"
        k_f2 = f"ng_input_25kg_{st.session_state.ng_reset_token}"

        st.markdown(f"<h4 style='color: {th['text']};'>📥 Manual Bag Quantity Inputs</h4>", unsafe_allow_html=True)
        col_inp1, col_inp2 = st.columns(2)

        with col_inp1:
            bags_50 = st.number_input("📦 Category A (50 Kg Bags)", min_value=0, max_value=1000000, step=10, value=0, key=k_f1)
        with col_inp2:
            bags_25 = st.number_input("📦 Category B (25 Kg Bags)", min_value=0, max_value=1000000, step=10, value=0, key=k_f2)

        total_bags_combined = int(bags_50) + int(bags_25)

        w1_m_opt1 = bags_50 * 50.120 + bags_25 * 25.120
        mt_m_opt1 = w1_m_opt1 / 1000.0

        w1_m_opt2 = bags_50 * 50.0 + bags_25 * 25.0
        mt_m_opt2 = w1_m_opt2 / 1000.0

        st.markdown("<br>", unsafe_allow_html=True)
        res_col1, res_col2 = st.columns(2)

        with res_col1:
            st.markdown(f"""
                <div class="nextgen-card">
                    <div class="card-header">🔹 Option 1 (Precision Scale: 50.12 & 25.12)</div>
                    <div class="sub-metric">Total Volume: <b style="color: {th['text']};">{total_bags_combined:,} Bags</b></div>
                    <hr style="border-color: {th['border']}; margin: 10px 0; opacity: 0.4;">
                    <div class="sub-metric">Combined Net Weight:</div>
                    <div class="metric-value">{w1_m_opt1:,.3f} <span style="font-size: 14px; color: {th['subtext']};">Kgs</span></div>
                    <div style="color: #34d399; font-weight: 700; font-size: 15px; margin-top: 4px;">🚀 {mt_m_opt1:,.3f} Metric Tons (MT)</div>
                </div>
            """, unsafe_allow_html=True)

        with res_col2:
            st.markdown(f"""
                <div class="nextgen-card">
                    <div class="card-header">🔹 Option 2 (Standard Slabs: 50.0 & 25.0)</div>
                    <div class="sub-metric">Total Volume: <b style="color: {th['text']};">{total_bags_combined:,} Bags</b></div>
                    <hr style="border-color: {th['border']}; margin: 10px 0; opacity: 0.4;">
                    <div class="sub-metric">Combined Net Weight:</div>
                    <div class="metric-value">{w1_m_opt2:,.2f} <span style="font-size: 14px; color: {th['subtext']};">Kgs</span></div>
                    <div style="color: #34d399; font-weight: 700; font-size: 15px; margin-top: 4px;">🚀 {mt_m_opt2:,.3f} Metric Tons (MT)</div>
                </div>
            """, unsafe_allow_html=True)

        if st.button("🔄 Reset Manual Calculator State", type="secondary"):
            st.session_state.ng_reset_token += 1
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    st.dataframe(working_df, use_container_width=True)

    # ==========================================================================
    # ACTION BUTTONS: HTML PRINT VIEW, EXCEL DOWNLOAD & EMAIL DISPATCH
    # ==========================================================================
    st.markdown("---")
    st.markdown("### 🚀 Export, Print & Email Options")
    
    col_act1, col_act2, col_act3 = st.columns(3)

    with col_act1:
        html_table_string = working_df.to_html(classes='table table-striped', index=False, border=0)
        print_html_code = f"""
            <html>
                <head>
                    <title>Tonnage Report Print View</title>
                    <style>
                        body {{ font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; color: #1e293b; }}
                        h2 {{ color: #0f172a; margin-bottom: 5px; }}
                        p {{ color: #475569; font-size: 13px; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 11px; table-layout: auto; }}
                        th, td {{ border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; word-wrap: break-word; }}
                        th {{ background-color: #1e293b; color: white; font-weight: 600; }}
                        tr:nth-child(even) {{ background-color: #f8fafc; }}
                        @media print {{
                            body {{ padding: 0; }}
                            table {{ page-break-inside: auto; }}
                            tr {{ page-break-inside: avoid; page-break-after: auto; }}
                        }}
                    </style>
                </head>
                <body>
                    <h2>Enterprise Vehicle Tonnage Summary Report</h2>
                    <p><b>Generated on:</b> {get_ist_now().strftime('%d-%m-%Y %H:%M:%S IST')}</p>
                    {html_table_string}
                    <script>
                        window.onload = function() {{ window.print(); }}
                    </script>
                </body>
            </html>
        """
        encoded_html = io.BytesIO(print_html_code.encode('utf-8'))
        st.download_button(
            "🖨️ Print / Save Table as PDF",
            data=encoded_html,
            file_name=f"Printable_Tonnage_Report_{get_ist_now().strftime('%Y%m%d_%H%M%S')}.html",
            mime="text/html",
            type="secondary",
            help="Downloads a responsive formatted HTML view that opens cleanly in print layout."
        )

    with col_act2:
        excel_buffer = io.BytesIO()
        working_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        st.download_button(
            "📥 Download Filtered Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"Tonnage_Report_{get_ist_now().strftime('%Y-%m-%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    with col_act3:
        with st.expander("✉️ Send Report via Email (HTML Summary & Attachment)"):
            with st.form("email_dispatch_form"):
                st.markdown("**SMTP Server Settings (Gmail / Corporate Hub):**")
                smtp_host = st.text_input("SMTP Server", "smtp.gmail.com")
                smtp_port = st.number_input("SMTP Port", value=587, step=1)
                sender_email = st.text_input("Sender Email (From)", "your_email@gmail.com")
                sender_pass = st.text_input("Sender App Password", type="password", help="Use Gmail App Password")
                
                st.markdown("---")
                email_to = st.text_input("Recipient Email(s) separated by comma (,):", "recipient@example.com")
                email_sub = st.text_input("Email Subject", "🚨 Executive Vehicle Tonnage Summary Report")
                
                default_mail_body = (
                    "Dear Leadership / Management Team,\n\n"
                    "Please find below the summary table and attached Excel report containing vehicle tonnage details.\n\n"
                    "Best Regards,\n"
                    "Logistics Hub"
                )
                email_body = st.text_area("Email Message", default_mail_body, height=120)
                
                send_email_btn = st.form_submit_button("📨 Send Professional Email", type="primary")

                if send_email_btn:
                    recipients = [e.strip() for e in email_to.split(",") if e.strip()]
                    if not recipients:
                        st.error("❌ Kripya kam se kam ek valid recipient email enter karein.")
                    else:
                        try:
                            msg = EmailMessage()
                            msg['Subject'] = email_sub
                            msg['From'] = sender_email
                            msg['To'] = ", ".join(recipients)
                            
                            table_html_snippet = working_df.head(50).to_html(index=False, border=1, classes='styled-table')
                            
                            html_content = f"""
                            <html>
                            <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; line-height: 1.6; padding: 15px;">
                                <div style="background: #1e293b; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                                    <h2 style="margin: 0; color: #38bdf8;">🚚 Vehicle Tonnage Summary Report</h2>
                                    <p style="margin: 5px 0 0 0; font-size: 13px; color: #94a3b8;">Automated Report | Generated on {get_ist_now().strftime('%d-%m-%Y %H:%M:%S IST')}</p>
                                </div>
                                <p style="font-size: 14px; white-space: pre-wrap;">{email_body}</p>
                                <h3 style="color: #0f172a; border-bottom: 2px solid #cbd5e1; padding-bottom: 5px;">📋 Data Snapshot Preview (Top 50 Rows)</h3>
                                <style>
                                    .styled-table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
                                    .styled-table th {{ background-color: #1e293b; color: white; padding: 8px; text-align: left; border: 1px solid #cbd5e1; }}
                                    .styled-table td {{ padding: 7px 8px; border: 1px solid #cbd5e1; text-align: left; color: #334155; }}
                                    .styled-table tr:nth-child(even) {{ background-color: #f8fafc; }}
                                </style>
                                {table_html_snippet}
                                <p style="margin-top: 25px; font-size: 12px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px;">
                                    <i>Note: The complete detailed dataset has been attached as an Excel sheet (`.xlsx`) to this email.</i><br>
                                    <b>Enterprise Logistics Hub</b>
                                </p>
                            </body>
                            </html>
                            """
                            msg.set_content(email_body)
                            msg.add_alternative(html_content, subtype='html')

                            excel_data = io.BytesIO()
                            working_df.to_excel(excel_data, index=False)
                            excel_data.seek(0)
                            msg.add_attachment(
                                excel_data.getvalue(),
                                maintype='application',
                                subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                filename=f"Tonnage_Report_{get_ist_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            )

                            server = smtplib.SMTP(smtp_host, int(smtp_port))
                            server.starttls()
                            server.login(sender_email, sender_pass)
                            server.send_message(msg)
                            server.quit()

                            st.success(f"✅ Professional email with Excel attachment successfully dispatched to: **{', '.join(recipients)}**!")
                        except Exception as mail_err:
                            st.error(f"❌ Email sending failed. Error details: {str(mail_err)}")
else:
    st.info("ℹ️ Kripya left sidebar se apni billing export file upload karein ya saved table select karein.")
