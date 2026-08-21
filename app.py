import streamlit as st
import pandas as pd
import openpyxl
import datetime
import io
import re

# Page Configuration
st.set_page_config(page_title="Sales Order Automation Hub", page_icon="🚀", layout="centered")

st.markdown("""
    <style>
        #GithubIcon { visibility: hidden; }
        .stButton>button { width: 100%; background-color: #10b981 !important; color: #ffffff !important; font-size: 16px; font-weight: 700; padding: 14px; border-radius: 8px; border: none; }
        .stButton>button:hover { background-color: #059669 !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Sales Order Automation Hub")

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

uploaded_inputs = st.file_uploader("Upload Multiple Demand Excel Files", type=["xlsx", "xls"], accept_multiple_files=True, key="inputs")

if st.button("🚀 Process Batch Orders", type="primary"):
    if uploaded_inputs:
        st.session_state.processed_files = []
        with st.spinner("⚡ Processing..."):
            try:
                with open("Output.xlsx", "rb") as f:
                    template_bytes = f.read()
                
                today_date = datetime.date.today().strftime("%Y-%m-%d")
                timestamp = datetime.datetime.now().strftime("%H%M%S")

                for uploaded_file in uploaded_inputs:
                    df_input = pd.read_excel(io.BytesIO(uploaded_file.getvalue()), header=None)

                    # 1. Logic: Find FG
                    fg_row, fg_col = -1, -1
                    for r in range(df_input.shape[0]):
                        for c in range(df_input.shape[1]):
                            if str(df_input.iloc[r, c]).strip().upper().startswith("FG"):
                                fg_row, fg_col = r, c
                                break
                        if fg_row != -1: break
                    if fg_row == -1: continue

                    # 2. Total Col Detection
                    total_col = df_input.shape[1]
                    for cSearch in range(fg_col, df_input.shape[1]):
                        is_total = False
                        for scan_r in range(max(0, fg_row - 10), min(fg_row + 3, df_input.shape[0])):
                            cell_val = str(df_input.iloc[scan_r, cSearch]).strip().upper()
                            if any(kw in cell_val for kw in ["TOTAL", "SUM", "TOTA", "TOT", "TTL", "NET"]):
                                is_total = True; break
                        if is_total: total_col = cSearch; break

                    # 3. Route & Agency Detection
                    route_num = "22"
                    for r in range(fg_row):
                        for c in range(min(total_col, 30)):
                            cell_val = str(df_input.iloc[r, c]).strip()
                            if 1 <= len(cell_val) <= 5 and any(char.isdigit() for char in cell_val):
                                route_num = cell_val; break
                        if route_num != "22": break
                    
                    agency_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        valid_c = 0
                        for rCheck in range(fg_row + 1, df_input.shape[0]):
                            v = str(df_input.iloc[rCheck, cSearch]).replace('.0', '').strip()
                            if v.isdigit() and 1 <= len(v) <= 5: valid_c += 1
                        if valid_c > 0: agency_col = cSearch; break

                    dr_code_col = -1
                    for cSearch in range(fg_col - 1, -1, -1):
                        for offset in range(1, 5):
                            if fg_row+offset < df_input.shape[0] and re.match(r'^DR\d+', str(df_input.iloc[fg_row+offset, cSearch]).upper()):
                                dr_code_col = cSearch; break
                        if dr_code_col != -1: break

                    valid_cols = [(c, str(df_input.iloc[fg_row, c])) for c in range(fg_col, total_col) if "TOTAL" not in str(df_input.iloc[fg_row, c]).upper()]

                    # 4. Fill Template (No Header Modification)
                    wb_valid = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_valid = wb_valid["Order Data"]
                    wb_missing = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws_missing = wb_missing["Order Data"]

                    valid_row, missing_row = 6, 6
                    
                    for r in range(fg_row + 1, df_input.shape[0]):
                        agency = df_input.iloc[r, agency_col] if agency_col >= 0 else None
                        if pd.notna(agency) and str(agency).strip().isdigit():
                            agency_val = int(str(agency).replace('.0', ''))
                            has_dr = False
                            clean_dr = str(df_input.iloc[r, dr_code_col]).replace('.0', '') if dr_code_col >= 0 else ""
                            if clean_dr != "nan" and clean_dr != "": has_dr = True
                            
                            target_ws = ws_valid if has_dr else ws_missing
                            current_r = valid_row if has_dr else missing_row
                            
                            for c, fg_code in valid_cols:
                                qty = df_input.iloc[r, c]
                                if pd.notna(qty) and float(qty) > 0:
                                    target_ws.cell(row=current_r, column=2, value=1) # Simplified Order logic
                                    target_ws.cell(row=current_r, column=7, value=clean_dr if has_dr else f"NEW_{agency_val}")
                                    target_ws.cell(row=current_r, column=16, value=fg_code if str(fg_code).startswith("FG") else "FG500014")
                                    target_ws.cell(row=current_r, column=19, value=float(qty))
                                    current_r += 1
                            
                            if has_dr: valid_row = current_r
                            else: missing_row = current_r

                    # Save buffers
                    for wb, name_suffix in [(wb_valid, "_Valid"), (wb_missing, "_Missing")]:
                        buf = io.BytesIO()
                        wb.save(buf)
                        st.session_state.processed_files.append({"name": name_suffix, "data": buf.getvalue(), "filename": f"Output_{today_date}{name_suffix}.xlsx"})
                
                st.success("✅ Done!")
            except Exception as e: st.error(f"Error: {e}")
