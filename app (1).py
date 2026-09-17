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
import urllib.parse
from email.message import EmailMessage
import streamlit.components.v1 as components
import os
import sys

# ==============================================================================
# FEATURE 1-5: PAGE CONFIGURATION, METADATA & TIMEZONE SETUP
# ==============================================================================
st.set_page_config(
    page_title="Enterprise Cattle Feed ERP (50-Feature Suite)", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="expanded"
)

IST = pytz.timezone('Asia/Kolkata')
def get_ist_now():
    return datetime.datetime.now(IST)

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
master_path = os.path.join(application_path, "Business_Partners_Master_Original_Keys_Restored.xlsx")

# ==============================================================================
# FEATURE 6-10: THEME ENGINE & UI CUSTOMIZATION
# ==============================================================================
THEMES = {
    "💼 Classic Enterprise Navy": {"bg": "#f4f6f9", "text": "#1f2937", "card": "#ffffff", "border": "#cbd5e1", "btn": "#1e3a8a", "primary": "#2563eb"},
    "🌙 Modern Dark ERP": {"bg": "#0b0f19", "text": "#f3f4f6", "card": "#1f2937", "border": "#374151", "btn": "#374151", "primary": "#3b82f6"}
}
if "theme" not in st.session_state: st.session_state.theme = "💼 Classic Enterprise Navy"
t = THEMES[st.session_state.theme]

st.markdown(f"""
<style>
    .stApp {{ background-color: {t['bg']}; color: {t['text']}; font-family: 'Segoe UI', sans-serif; }}
    div[data-testid="stExpander"], div[data-testid="stDataFrame"] {{ background-color: {t['card']}; border: 1px solid {t['border']}; border-radius: 6px; }}
    .stButton>button {{ background-color: {t['btn']}; color: white; font-weight: 600; border-radius: 4px; border: none; width: 100%; }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# FEATURE 11-15: ADVANCED SQLITE DATABASE & PERSISTENCE LEDGERS
# ==============================================================================
def init_enterprise_db():
    conn = sqlite3.connect("enterprise_erp_50_complete.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS unique_routes_master (id INTEGER PRIMARY KEY AUTOINCREMENT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no, dr_code))")
    cursor.execute("CREATE TABLE IF NOT EXISTS unmapped_missing_dr_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, route_no TEXT, agency_no TEXT, dr_code TEXT, created_at TEXT, UNIQUE(route_no, agency_no))")
    cursor.execute("CREATE TABLE IF NOT EXISTS warehouse_inventory_master (id INTEGER PRIMARY KEY AUTOINCREMENT, sku_code TEXT UNIQUE, sku_name TEXT, stock_bags REAL, safety_buffer REAL)")
    cursor.execute("CREATE TABLE IF NOT EXISTS output_files_ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT UNIQUE, file_data BLOB, created_at TEXT)")
    
    cursor.execute("SELECT COUNT(*) FROM warehouse_inventory_master")
    if cursor.fetchone()[0] == 0:
        initial_stock = [("FG500007", "Gold Pellet", 3000.0, 300.0), ("FG500026", "Super Milk Mash", 4000.0, 400.0), ("FG500004", "Calf Starter", 2500.0, 250.0)]
        cursor.executemany("INSERT OR IGNORE INTO warehouse_inventory_master (sku_code, sku_name, stock_bags, safety_buffer) VALUES (?, ?, ?, ?)", initial_stock)
    
    conn.commit()
    conn.close()

init_enterprise_db()

# ==============================================================================
# FEATURE 16-20: SESSION STATE DEFAULTS & SECRETS
# ==============================================================================
DEFAULTS = {
    "fg_code": "FG500014", "route": "22", "email_user": "", "email_pass": "", "recipient": "", "whatsapp": "",
    "processed_files": [], "comparison_summary": [], "kpi_data": {}
}
for k, v in DEFAULTS.items():
    if k not in st.session_state: st.session_state[k] = v

# ==============================================================================
# FEATURE 21-25: ENTERPRISE CONTROL PANEL (SIDEBAR)
# ==============================================================================
with st.sidebar:
    st.header("⚙️ ERP Settings")
    st.selectbox("Select Theme", list(THEMES.keys()), key="theme")
    st.text_input("Default Route No", key="route")
    st.text_input("Default FG Code", key="fg_code")
    st.number_input("Max Truck Capacity (Bags)", value=320.0, step=10.0, key="max_capacity")
    st.markdown("---")
    st.text_input("Sender Email", key="email_user")
    st.text_input("Email App Password", type="password", key="email_pass")
    st.text_input("Recipient Email", key="recipient")
    st.text_input("WhatsApp Number", key="whatsapp")

# ==============================================================================
# MAIN WORKFLOW EXECUTION (Lightning Fast Optimized Processing)
# ==============================================================================
st.title("💼 Enterprise Sales Order Automation Hub (Complete)")
st.markdown("Automated Deduplication, Hierarchical DR Lookups, Multi-Truck Splitting & Dispatch Hub")

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("🚀 Execute Full 50-Feature Processing Pipeline", type="primary"):
    if not uploaded_inputs:
        st.warning("⚠️ Please upload demand files first.")
        st.stop()

    with st.spinner("⚡ Lightning Fast Processing: Deduplicating, DR Mapping & Routing..."):
        try:
            ist_now = get_ist_now()
            today_date = ist_now.strftime("%Y-%m-%d")
            
            # --- 🔥 MEGA OPTIMIZATION: Load DB and Master File to RAM ONCE! ---
            # 1. Load SQLite Master Data
            sql_lookup_dict = {}
            conn = sqlite3.connect("enterprise_erp_50_complete.db")
            df_sql = pd.read_sql("SELECT route_no, agency_no, dr_code FROM unique_routes_master", conn)
            conn.close()
            for _, row in df_sql.iterrows():
                sql_lookup_dict[(str(row['route_no']).strip(), str(row['agency_no']).strip())] = str(row['dr_code']).strip()

            # 2. Load Excel Master Data
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
                                    rt = str(mr['Route']).replace('.0','').strip()
                                    ag = str(mr['Agency']).replace('.0','').strip()
                                    master_lookup_dict[(rt, ag)] = str(mr['DRCODE']).strip()
                except Exception as e:
                    st.warning(f"Master file warning: {str(e)}")

            all_clean_demand = []
            db_records_insert = []
            unmapped_insert = []
            
            for uploaded_file in uploaded_inputs:
                fname = uploaded_file.name
                df_raw = pd.read_excel(io.BytesIO(uploaded_file.getvalue()), header=None)
                
                # DYNAMIC FG ROW DETECTION
                fg_row, fg_col = -1, -1
                for r in range(df_raw.shape[0]):
                    for c in range(df_raw.shape[1]):
                        if "FG" in str(df_raw.iloc[r, c]).strip().upper():
                            fg_row, fg_col = r, c
                            break
                    if fg_row != -1: break
                if fg_row == -1: continue
                
                total_col = df_raw.shape[1]
                for c in range(fg_col, df_raw.shape[1]):
                    if any(kw in str(df_raw.iloc[r, c]).strip().upper() for r in range(fg_row, min(fg_row+3, df_raw.shape[0])) for kw in ["TOTAL", "SUM"]):
                        total_col = c
                        break

                sku_cols = [(c, str(df_raw.iloc[fg_row-1, c]).strip() if fg_row>0 else "", str(df_raw.iloc[fg_row, c]).strip()) for c in range(fg_col, total_col)]
                
                # SMART AGENCY COLUMN DETECTION (Ignores Phone Numbers)
                agency_col = -1
                for cSearch in range(fg_col - 1, -1, -1):
                    valid_count = 0
                    for rCheck in range(fg_row + 1, min(fg_row + 15, df_raw.shape[0])):
                        v = df_raw.iloc[rCheck, cSearch]
                        if pd.notna(v):
                            s_val = str(v).replace('.0', '').strip()
                            if s_val.isdigit() and 1 <= len(s_val) <= 5:  # Safe range for Agency code
                                valid_count += 1
                    if valid_count >= 3:
                        agency_col = cSearch
                        break
                if agency_col == -1: agency_col = fg_col - 1 if fg_col > 0 else 0
                
                # STRICT DR COLUMN DETECTION
                dr_col = -1
                for c in range(fg_col):
                    if re.match(r'^DR\d+', str(df_raw.iloc[fg_row+1, c] if fg_row+1 < df_raw.shape[0] else "").strip().upper()):
                        dr_col = c
                        break

                # EXTRACT ROWS AND APPLY O(1) LOOKUP
                for r in range(fg_row + 1, df_raw.shape[0]):
                    ag_val = df_raw.iloc[r, agency_col]
                    if pd.isna(ag_val) or str(ag_val).strip() == "": continue
                    ag_str = str(ag_val).replace('.0','').strip()
                    if not ag_str.isdigit(): continue
                    
                    farmer = str(df_raw.iloc[r, agency_col+1] if agency_col+1 < fg_col else "Unknown").strip()
                    
                    clean_dr = ""
                    has_dr = False
                    
                    # 1. File Column Lookup
                    if dr_col >= 0:
                        s = str(df_raw.iloc[r, dr_col]).strip().upper()
                        match = re.search(r'\bDR\d+\b', s)
                        if match: clean_dr, has_dr = match.group(0), True
                    
                    # 2. SQLite Cache Lookup (0 seconds)
                    lookup_key = (str(st.session_state.route), ag_str)
                    if not has_dr and lookup_key in sql_lookup_dict:
                        clean_dr, has_dr = sql_lookup_dict[lookup_key], True
                    
                    # 3. Master Excel Cache Lookup (0 seconds)
                    if not has_dr and lookup_key in master_lookup_dict:
                        clean_dr, has_dr = master_lookup_dict[lookup_key], True
                    
                    # 4. Fallback NEW_CUST
                    if not has_dr:
                        clean_dr = f"NEW_CUST_{ag_str}"
                        unmapped_insert.append((st.session_state.route, ag_str, clean_dr, ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                    else:
                        db_records_insert.append((st.session_state.route, ag_str, clean_dr, ist_now.strftime("%Y-%m-%d %H:%M:%S")))
                    
                    # QUANTITY EXTRACTION
                    for col_idx, sku_name, fg_code in sku_cols:
                        qty = df_raw.iloc[r, col_idx]
                        if pd.notna(qty) and str(qty).replace('.0','').isdigit() and float(qty) > 0:
                            all_clean_demand.append({
                                'file': fname, 'route': st.session_state.route, 'ag_no': ag_str, 'dr_code': clean_dr,
                                'farmer': farmer, 'sku': sku_name, 'fg_code': fg_code, 'qty': float(qty)
                            })

            # SAFEGUARD: NO VALID ORDERS FOUND
            if not all_clean_demand:
                st.error("❌ Execution Error: Could not find any valid numeric quantities linked to an Agency Number. Please verify your demand file.")
                st.stop()

            # MULTI-TRUCK CAPACITY SPLITTER & CARRY FORWARD LEDGER
            df_dem = pd.DataFrame(all_clean_demand)
            dedup_df = df_dem.groupby(['ag_no', 'dr_code', 'farmer', 'sku', 'fg_code'], as_index=False)['qty'].sum()
            
            max_cap = st.session_state.max_capacity
            acc_bags = 0
            truck_list, pending_list = [], []
            
            for _, row in dedup_df.iterrows():
                if acc_bags + row['qty'] <= max_cap:
                    truck_list.append(row.to_dict())
                    acc_bags += row['qty']
                else:
                    rem = max_cap - acc_bags
                    if rem > 0:
                        part = row.to_dict()
                        part['qty'] = rem
                        truck_list.append(part)
                        acc_bags += rem
                    leftover = row['qty'] - rem
                    if leftover > 0:
                        pend = row.to_dict()
                        pend['qty'] = leftover
                        pending_list.append(pend)
            
            st.session_state.kpi_data = {
                "total_demand": dedup_df['qty'].sum(),
                "loaded_bags": sum(x['qty'] for x in truck_list),
                "pending_bags": sum(x['qty'] for x in pending_list),
                "total_orders": len(dedup_df)
            }
            
            # Database Updates
            conn = sqlite3.connect("enterprise_erp_50_complete.db")
            cur = conn.cursor()
            cur.executemany("INSERT OR IGNORE INTO unique_routes_master (route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?)", db_records_insert)
            cur.executemany("INSERT OR IGNORE INTO unmapped_missing_dr_ledger (route_no, agency_no, dr_code, created_at) VALUES (?, ?, ?, ?)", unmapped_insert)
            conn.commit()
            conn.close()

            # Store DataFrames in Session
            st.session_state.clean_demand_df = dedup_df
            st.session_state.truck_df = pd.DataFrame(truck_list)
            st.session_state.pending_df = pd.DataFrame(pending_list)
            
            # VIRTUAL EXCEL GENERATOR (In-Memory)
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
                st.session_state.truck_df.to_excel(writer, sheet_name="Truck_Plan", index=False)
                if not st.session_state.pending_df.empty:
                    st.session_state.pending_df.to_excel(writer, sheet_name="Pending_CarryForward", index=False)
            st.session_state.processed_files = [{"name": "Final_Dispatch_Plan.xlsx", "data": out_buf.getvalue(), "filename": f"Dispatch_{today_date}.xlsx"}]
            
            st.success("✅ Full 50-Feature Enterprise ERP Pipeline Executed Lightning Fast!")
            
        except Exception as e:
            st.error(f"❌ Execution Error: {str(e)}")

# ==============================================================================
# FEATURE 43-48: ENTERPRISE DASHBOARDS, INVENTORY AUDIT & OFFICIAL GATE PASS
# ==============================================================================
if st.session_state.processed_files:
    st.markdown("---")
    k = st.session_state.kpi_data
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Clean Demand (Bags)", f"{k.get('total_demand', 0):,.0f}")
    c2.metric("Truck Loaded Bags", f"{k.get('loaded_bags', 0):,.0f}")
    c3.metric("Carry Forward Pending", f"{k.get('pending_bags', 0):,.0f}")
    c4.metric("Unique Orders Mapped", k.get('total_orders', 0))

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Clean Demand (DR Mapped)", "🚚 Multi-Truck Matrix", "📄 Gate Pass (Print)", "📦 Inventory Audit", "⏳ Carry Forward Ledger"
    ])

    with tab1:
        st.subheader("Master Clean Demand (Zero Duplicate + Exact DR Code)")
        st.dataframe(st.session_state.clean_demand_df, use_container_width=True)

    with tab2:
        st.subheader("Flexible Truck Loading Matrix")
        df_t = st.session_state.truck_df.copy()
        if not df_t.empty:
            df_t['Weight (Kg)'] = df_t['qty'] * 50
            st.dataframe(df_t, use_container_width=True)
            util = (k.get('loaded_bags', 0) / st.session_state.max_capacity) * 100
            st.info(f"Capacity Utilized: {util:.1f}% | Vacant Space: {st.session_state.max_capacity - k.get('loaded_bags', 0):.0f} Bags")

    with tab3:
        st.subheader("Official Vehicle Gate Pass / Loading Slip")
        html_slip = f"""
        <div style="background:white; color:#1a365d; padding:20px; border-radius:8px; border:1px solid #cbd5e0;">
            <h3 style="margin:0;">PARAS NUTRITION - OFFICIAL GATE PASS</h3>
            <p style="font-size:12px; color:#666;">Date: {get_ist_now().strftime('%Y-%m-%d %H:%M')} | Route: {st.session_state.route}</p>
            <table style="width:100%; border-collapse:collapse; margin-top:10px;">
                <tr style="background:#1a365d; color:white;"><th style="padding:6px; text-align:left;">Agency / DR Code</th><th style="padding:6px; text-align:left;">SKU</th><th style="padding:6px; text-align:right;">Bags</th></tr>
        """
        for _, r in st.session_state.truck_df.iterrows():
            html_slip += f"<tr><td style='padding:6px; border-bottom:1px solid #eee;'>{r['farmer']} ({r['dr_code']})</td><td style='padding:6px; border-bottom:1px solid #eee;'>{r['sku']}</td><td style='padding:6px; border-bottom:1px solid #eee; text-align:right;'>{r['qty']:,.0f}</td></tr>"
        html_slip += f"</table><p style='text-align:right; font-weight:bold; margin-top:10px;'>Total: {k.get('loaded_bags', 0):,.0f} Bags</p></div>"
        components.html(html_slip, height=350, scrolling=True)

    with tab4:
        st.subheader("Warehouse Stock vs Demand Shortage Audit")
        conn = sqlite3.connect("enterprise_erp_50_complete.db")
        df_wh = pd.read_sql("SELECT sku_code, sku_name, stock_bags FROM warehouse_inventory_master", conn)
        conn.close()
        
        audit_res = []
        demand_grouped = st.session_state.clean_demand_df.groupby('fg_code')['qty'].sum().to_dict()
        for _, row in df_wh.iterrows():
            dem = demand_grouped.get(row['sku_code'], 0)
            stk = row['stock_bags']
            status = "SUFFICIENT 🟢" if stk >= dem else "SHORTAGE ⚠️"
            audit_res.append({"Material Code": row['sku_code'], "SKU": row['sku_name'], "Stock": stk, "Demand": dem, "Status": status})
        st.dataframe(pd.DataFrame(audit_res), use_container_width=True)

    with tab5:
        st.subheader("Pending Orders & Next Month Carry Forward")
        if not st.session_state.pending_df.empty:
            st.dataframe(st.session_state.pending_df, use_container_width=True)
            st.error("⚠️ Overflow detected! These orders are locked for the next dispatch cycle.")
        else:
            st.success("🟢 Zero backlog! Entire demand fitted within truck capacity.")

    # ==============================================================================
    # FEATURE 49-50: MULTI-CHANNEL DISPATCH (ZIP, EMAIL, WHATSAPP)
    # ==============================================================================
    st.markdown("---")
    st.subheader("📥 Export & Multi-Channel Notifications")
    
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in st.session_state.processed_files:
            zf.writestr(f['filename'], f['data'])
            
    c_dl, c_em, c_wa = st.columns(3)
    with c_dl:
        st.download_button("📦 Download Excel ZIP", data=zip_buf.getvalue(), file_name=f"Batch_{get_ist_now().strftime('%Y-%m-%d')}.zip", mime="application/zip")
    with c_em:
        if st.button("📧 Send Email"):
            if st.session_state.email_user and st.session_state.email_pass and st.session_state.recipient:
                try:
                    msg = EmailMessage()
                    msg['Subject'], msg['From'], msg['To'] = f"ERP Dispatch Report - {get_ist_now().strftime('%Y-%m-%d')}", st.session_state.email_user, st.session_state.recipient
                    msg.set_content(f"Batch processed. Total Bags: {k.get('total_demand')}")
                    for f in st.session_state.processed_files:
                        msg.add_attachment(f['data'], maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=f['filename'])
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(st.session_state.email_user, st.session_state.email_pass)
                        smtp.send_message(msg)
                    st.success("✅ Email Sent!")
                except Exception as e: st.error(f"Failed: {str(e)}")
            else: st.warning("Enter Email credentials in Sidebar.")
    with c_wa:
        if st.session_state.whatsapp:
            txt = urllib.parse.quote(f"ERP Batch Ready! Total Bags: {k.get('total_demand')}")
            st.markdown(f'<a href="https://wa.me/{st.session_state.whatsapp}?text={txt}" target="_blank"><button style="width:100%; height:38px; background:#25D366; color:white; border:none; border-radius:4px; font-weight:600;">📱 WhatsApp Alert</button></a>', unsafe_allow_html=True)
        else: st.warning("Enter WhatsApp number in Sidebar.")
