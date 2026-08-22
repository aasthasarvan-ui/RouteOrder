import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io
import re
import zipfile
import smtplib
import urllib.parse
from email.message import EmailMessage

# Page Configuration & Styling
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
        
        .stButton>button {
            width: 100%;
            background-color: #10b981 !important;
            color: #ffffff !important;
            font-size: 16px;
            font-weight: 700;
            padding: 14px;
            border-radius: 8px;
            border: none;
        }
        .stButton>button:hover {
            background-color: #059669 !important;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar Settings & Dynamic Mapping
st.sidebar.title("⚙️ System Settings")
default_fg_code = st.sidebar.text_input("Default Fallback FG Code", value="FG500014")

col_mapping_input = st.sidebar.text_area(
    "Direct Column Index Mapping (ColIndex:Code)", 
    value="36:FG500014AJ\n37:FG500014AK",
    help="Excel file ke exact column index ke anusaar code assign karein jahan header blank ho."
)

st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Agency-wise FG Override")
agency_fg_override = st.sidebar.text_area(
    "Agency:CustomFG (e.g., 101:FG500014N01)", 
    value="101:FG500014N01\n102:FG500014N02",
    help="Agar kisi specific agency ke paas FG code nahi hai, toh yahan unka custom code define karein."
)

default_fallback_route = st.sidebar.text_input("Default Route Fallback", value="22")

direct_col_mapping = {}
for line in col_mapping_input.split('\n'):
    if ':' in line:
        parts = line.split(':')
        idx_str = parts[0].strip()
        if idx_str.isdigit():
            direct_col_mapping[int(idx_str)] = parts[1].strip()

agency_override_map = {}
for line in agency_fg_override.split('\n'):
    if ':' in line:
        ag, fg = line.split(':')
        ag_clean = ag.strip()
        if ag_clean.isdigit():
            agency_override_map[int(ag_clean)] = fg.strip()

st.sidebar.markdown("---")
st.sidebar.subheader("📧 Email Dispatch Settings")

# --- Streamlit Secrets Integration ---
email_user = st.sidebar.text_input("Sender Email ID", value=st.secrets.get("email", {}).get("sender_email", ""))
email_pass = st.sidebar.text_input("Email App Password", type="password", value=st.secrets.get("email", {}).get("app_password", ""))
recipient_email = st.sidebar.text_input("Recipient Email", value=st.secrets.get("email", {}).get("recipient_email", ""))

st.sidebar.markdown("---")
st.sidebar.subheader("📱 WhatsApp Notification")
whatsapp_num = st.sidebar.text_input("WhatsApp Number (e.g., 919876543210)")

st.title("📊 Enterprise Sales Order Automation Hub")
st.markdown("Upload multiple **Inbound Demand Files** to process orders, apply direct column mappings, view KPIs, and export audit reports.")
st.markdown("---")

# Session State Initialization
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []
if 'comparison_summary' not in st.session_state:
    st.session_state.comparison_summary = []
if 'skipped_rows_log' not in st.session_state:
    st.session_state.skipped_rows_log = []
if 'history' not in st.session_state:
    st.session_state.history = []
if 'kpi_data' not in st.session_state:
    st.session_state.kpi_data = {"input_qty": 0, "gen_qty": 0, "valid_count": 0, "missing_count": 0, "skipped_count": 0}

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

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
                
                today_date = datetime.date.today().strftime("%Y-%m-%d")
                timestamp = datetime.datetime.now().strftime("%H%M%S")

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
                    route_num = default_fallback_route
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
                            
                            if cell_val != "" and 1 <= len(cell_val) <= 5:
                                if any(char.isdigit() for char in cell_val):
                                    route_num = cell_val
                                    break
                        if route_num != default_fallback_route:
                            break

                    safe_route_num = "".join(c if c.isalnum() or c in ('-', '_') else "-" for c in str(route_num))

                    # 4. Smart Agency Detection (with robust fallback)
                    agency_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        valid_count = 0
                        for rCheck in range(fg_row + 1, min(fg_row + 15, df_input.shape[0])):
                            v = df_input.iloc[rCheck, cSearch]
                            if pd.notna(v):
                                # Enhanced fuzzy extraction to support clean numeric IDs safely
                                cleaned_digits = re.sub(r'\D', '', str(v))
                                if 1 <= len(cleaned_digits) <= 5:
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
                        upper_fg = fg_code.upper()
                        
                        if any(kw in upper_fg for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                            break
                        
                        valid_cols.append((c, fg_code))

                    # 6. Load Template for Valid & Missing DR Orders separately
                    wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_valid = wb_valid["Order Data"] if "Order Data" in wb_valid.sheetnames else wb_valid.active

                    wb_missing = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_missing = wb_missing["Order Data"] if "Order Data" in wb_missing.sheetnames else wb_missing.active

                    valid_row = 6
                    missing_row = 6
                    valid_order_num = 1
                    missing_order_num = 1
                    
                    agency_counts_valid = {}
                    agency_counts_missing = {}
                    
                    valid_items_created = 0
                    missing_items_created = 0
                    
                    file_comparison_rows = []

                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        
                        if pd.isna(agency) or str(agency).strip() in ["", "nan", "None"]:
                            continue
                        
                        # Advanced fuzzy regex cleanup for agency string to prevent formatting issues
                        agency_digits = re.sub(r'\D', '', str(agency))
                        if not agency_digits or not (1 <= len(agency_digits) <= 5):
                            st.session_state.skipped_rows_log.append({
                                "File Name": short_filename,
                                "Row Index": r + 1,
                                "Agency Value": str(agency),
                                "Reason": "Invalid or Non-numeric Agency Number"
                            })
                            total_skipped_rows += 1
                            continue

                        agency_val = int(agency_digits)
                        
                        # --- ROBUST DR CODE & ZERO HANDLER ---
                        has_dr_code = False
                        clean_dr = ""
                        
                        if dr_code_col >= 0 and dr_code_col < df_input.shape[1]:
                            raw_dr = df_input.iloc[r, dr_code_col]
                            if pd.notna(raw_dr):
                                val_str = str(raw_dr).replace('.0', '').strip()
                                upper_str = val_str.upper()
                                if "DR" in upper_str and any(char.isdigit() for char in upper_str) and upper_str != "0":
                                    has_dr_code = True
                                    clean_dr = val_str

                        if not has_dr_code:
                            for c_scan in range(fg_col):
                                cell_val = df_input.iloc[r, c_scan]
                                if pd.notna(cell_val):
                                    val_str = str(cell_val).replace('.0', '').strip()
                                    upper_str = val_str.upper()
                                    if "DR" in upper_str and any(char.isdigit() for char in upper_str) and upper_str != "0":
                                        has_dr_code = True
                                        clean_dr = val_str
                                        break

                        # Check item quantities for this row
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

                        # Route based on DR Code presence
                        if has_dr_code:
                            agency_counts_valid[agency_val] = agency_counts_valid.get(agency_val, 0) + 1
                            current_seq = agency_counts_valid[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-{current_seq}"
                            
                            target_ws = ws_valid
                            current_r = valid_row
                            order_num = valid_order_num
                            dr_to_use = clean_dr
                            file_category = "Valid DR"
                        else:
                            agency_counts_missing[agency_val] = agency_counts_missing.get(agency_val, 0) + 1
                            current_seq = agency_counts_missing[agency_val]
                            ref_number = f"RT-{route_num}-{agency_val}-{today_date}-NEW" if current_seq == 1 else f"RT-{route_num}-{agency_val}-{today_date}-NEW-{current_seq}"
                            
                            target_ws = ws_missing
                            current_r = missing_row
                            order_num = missing_order_num
                            dr_to_use = f"NEW_CUST_{agency_val}"
                            file_category = "Missing DR"

                        item_id = 10
                        for c, fg_code, qty_val in valid_row_quantities:
                            cleaned_fg = str(fg_code).strip()
                            upper_fg = cleaned_fg.upper()
                            
                            # --- STRICT BLANK & DIRECT INDEX MAPPING LOGIC ---
                            if upper_fg.startswith("FG"):
                                current_fg = cleaned_fg
                            elif agency_val in agency_override_map:
                                current_fg = agency_override_map[agency_val]
                            elif upper_fg in ["", "NAN", "NONE"]:
                                if c in direct_col_mapping:
                                    current_fg = direct_col_mapping[c]
                                else:
                                    current_fg = default_fg_code
                            elif c in direct_col_mapping:
                                current_fg = direct_col_mapping[c]
                            else:
                                current_fg = default_fg_code
                            
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
                            valid_row = current_r
                            valid_order_num += 1
                            valid_items_created += 1
                            total_valid_orders += 1
                        else:
                            missing_row = current_r
                            missing_order_num += 1
                            missing_items_created += 1
                            total_missing_orders += 1

                    if valid_items_created > 0:
                        buf_valid = io.BytesIO()
                        wb_valid.properties.creator = "Microsoft Excel"
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
                        wb_missing.properties.creator = "Microsoft Excel"
                        wb_missing.save(buf_missing)
                        buf_missing.seek(0)
                        
                        st.session_state.processed_files.append({
                            "name": short_filename + " (Missing DR / New Customer)",
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

                # Store KPI metrics & History
                st.session_state.kpi_data = {
                    "input_qty": total_input_qty,
                    "gen_qty": total_gen_qty,
                    "valid_count": total_valid_orders,
                    "missing_count": total_missing_orders,
                    "skipped_count": total_skipped_rows
                }
                
                st.session_state.history.append({
                    "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Files Processed": len(uploaded_inputs),
                    "Total Qty": total_input_qty,
                    "Status": "Success"
                })

                st.success("✅ Batch Processing, Fixed Column Mapping & Audit Complete!")

            except Exception as e:
                st.error(f"❌ Error aagaya: {str(e)}")
    else:
        st.warning("⚠️ Kripya pehle demand files upload karein!")

# Display KPI Summary Cards
if st.session_state.processed_files or st.session_state.skipped_rows_log:
    st.markdown("---")
    st.markdown("### 📈 Batch Performance & KPI Summary")
    kpi = st.session_state.kpi_data
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Input Qty", f"{kpi['input_qty']:,.0f}")
    col2.metric("Generated Qty", f"{kpi['gen_qty']:,.0f}")
    col3.metric("Valid Orders", kpi['valid_count'])
    col4.metric("Missing Orders", kpi['missing_count'])
    col5.metric("Skipped Rows", kpi['skipped_count'], delta_color="inverse")

    st.markdown("---")
    st.markdown("### 📥 Bulk Download & Notifications")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in st.session_state.processed_files:
            zip_file.writestr(item['filename'], item['data'])
    
    col_zip, col_pdf, col_email, col_wa = st.columns(4)
    
    with col_zip:
        st.download_button(
            label="📦 Download ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"Batch_Orders_ZIP_{datetime.date.today().strftime('%Y-%m-%d_%H%M%S')}.zip",
            mime="application/zip",
            key="zip_download"
        )
        
    with col_pdf:
        # --- PDF/Text Summary Report Feature ---
        summary_txt = f"""=== ENTERPRISE SALES ORDER SUMMARY ===
Date: {datetime.date.today()}
----------------------------------------
Total Input Quantity : {kpi['input_qty']:,.0f}
Total Generated Qty  : {kpi['gen_qty']:,.0f}
Valid DR Orders      : {kpi['valid_count']}
Missing DR Orders    : {kpi['missing_count']}
Skipped Rows Logged  : {kpi['skipped_count']}
----------------------------------------
Generated Files Count: {len(st.session_state.processed_files)}
Status: Successfully Processed & Audited
========================================"""
        st.download_button(
            label="📄 Download Summary Report",
            data=summary_txt.encode('utf-8'),
            file_name=f"Summary_Report_{datetime.date.today()}.txt",
            mime="text/plain",
            key="summary_txt_download"
        )
    
    with col_email:
        # --- HTML Rich Email Notification Feature ---
        if st.button("📧 Send HTML Email"):
            if email_user and email_pass and recipient_email:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = f"🚀 Sales Orders Batch Execution Report - {datetime.date.today()}"
                    msg['From'] = email_user
                    msg['To'] = recipient_email
                    
                    html_content = f"""
                    <html>
                      <body style="font-family: Arial, sans-serif; color: #333;">
                        <h2 style="color: #10b981;">📊 Sales Order Batch Automation Hub</h2>
                        <p>Hello Team,</p>
                        <p>The daily inbound demand batch has been processed successfully. Here are the key highlights:</p>
                        <table style="border-collapse: collapse; width: 100%; max-width: 500px;">
                          <tr style="background-color: #f3f4f6;"><th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Metric</th><th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Value</th></tr>
                          <tr><td style="border: 1px solid #ddd; padding: 8px;">Total Input Qty</td><td style="border: 1px solid #ddd; padding: 8px;"><b>{kpi['input_qty']:,.0f}</b></td></tr>
                          <tr><td style="border: 1px solid #ddd; padding: 8px;">Valid Orders</td><td style="border: 1px solid #ddd; padding: 8px;">{kpi['valid_count']}</td></tr>
                          <tr><td style="border: 1px solid #ddd; padding: 8px;">Missing DR Orders</td><td style="border: 1px solid #ddd; padding: 8px;">{kpi['missing_count']}</td></tr>
                          <tr><td style="border: 1px solid #ddd; padding: 8px;">Skipped Rows</td><td style="border: 1px solid #ddd; padding: 8px;">{kpi['skipped_count']}</td></tr>
                        </table>
                        <p style="margin-top: 20px;">Please find the generated order files attached herewith.</p>
                        <p style="color: #666; font-size: 12px;">Automated via Sales Order Hub</p>
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
                    
                    st.success("✅ HTML Rich Email dispatched successfully!")
                except Exception as e:
                    st.error(f"❌ Email failed: {str(e)}")
            else:
                st.warning("⚠️ Enter email credentials in sidebar!")

    with col_wa:
        if whatsapp_num:
            wa_text = f"Hello, Sales Order Batch Report is ready. Total Qty: {kpi['input_qty']}, Valid Orders: {kpi['valid_count']}, Missing Orders: {kpi['missing_count']}."
            encoded_wa = urllib.parse.quote(wa_text)
            wa_link = f"https://wa.me/{whatsapp_num}?text={encoded_wa}"
            st.markdown(f'<a href="{wa_link}" target="_blank" style="text-decoration:none;"><button style="width:100%; padding:14px; background:#25D366; color:white; border:none; border-radius:8px; font-weight:bold;">📱 Send WhatsApp</button></a>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("##### Individual File Downloads:")
    for i, item in enumerate(st.session_state.processed_files):
        st.success(f"✅ Processed: {item['name']} -> Orders created: {item['orders']}")
        if st.download_button(
            label=f"📥 Download {item['name']}",
            data=item['data'],
            file_name=item['filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_file_{i}_{item['filename']}"
        ):
            st.toast(f"🎉 '{item['filename']}' successfully download ho gaya hai!", icon="📥")

# Display Agency-wise Material & Quantity Breakdown with Comparison
if st.session_state.comparison_summary:
    st.markdown("---")
    st.markdown("### 📋 Agency-wise Material, Quantity & Input Comparison")
    st.markdown("Yeh table dikhati hai ki kis agency ne kis Material (FG Code) ki kitni quantity mangi aur generate hui:")
    
    combined_df = pd.concat(st.session_state.comparison_summary, ignore_index=True)
    agency_material_summary = combined_df.groupby(["Agency", "DR Code", "FG Code"], as_index=False).agg({
        "Input Qty": "sum",
        "Generated Qty": "sum"
    })
    agency_material_summary["Difference"] = agency_material_summary["Input Qty"] - agency_material_summary["Generated Qty"]
    
    st.dataframe(agency_material_summary, use_container_width=True)

# Display Skipped / Invalid Rows Exception Logger Table
if st.session_state.skipped_rows_log:
    st.markdown("---")
    st.markdown("### ⚠️ Skipped / Invalid Rows Exception Log")
    df_skipped = pd.DataFrame(st.session_state.skipped_rows_log)
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
        file_name=f"Audit_Reconciliation_Report_{datetime.date.today().strftime('%Y-%m-%d_%H%M%S')}.csv",
        mime="text/csv",
        key="audit_csv_download"
    )

    st.markdown("---")
    st.info(f"📊 **Batch Summary:** Total Output Files Generated: {len(st.session_state.processed_files)}")

# Session History Section
if st.session_state.history:
    st.markdown("---")
    with st.expander("🕒 View Session Processing History"):
        st.dataframe(pd.DataFrame(st.session_state.history), use_container_width=True)
