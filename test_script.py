import streamlit as st
import pandas as pd
import openpyxl
import os
import re
import io

st.set_page_config(page_title="Auto DRCODE Mapping Tool", layout="wide")
st.title("💼 Smart Input & Master DRCODE Mapping Hub (Format Preserved)")

uploaded_file = st.file_uploader("Upload Input Demand Excel File", type=["xlsx", "xls"])
master_file_input = st.file_uploader("Upload Master Route File (Optional - agar GitHub mein nahi hai)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 1. Load Master Route File (Dono options: Device upload ya repository file)
        master_df = None
        master_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"
        
        if master_file_input is not None:
            master_df = pd.read_excel(master_file_input, sheet_name="Master Data", header=2)
            st.success("Master file uploaded successfully from device!")
        elif os.path.exists(master_path):
            master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
            st.info("Master file loaded automatically from repository.")
        else:
            st.error("⚠️ Master file nahi mili! Kripya Master file upload karein.")
        
        if master_df is not None:
            master_df.columns = master_df.columns.astype(str).str.strip()
            
            # Build Master Mapping Dictionary: (Route, Agency) -> DRCODE
            mapping_dict = {}
            if 'Route' in master_df.columns and 'Agency' in master_df.columns and 'DRCODE' in master_df.columns:
                master_subset = master_df[['Route', 'Agency', 'DRCODE']].dropna(subset=['DRCODE'])
                for _, r_item in master_subset.iterrows():
                    rt_key = str(r_item['Route']).replace('.0', '').strip()
                    ag_key = str(r_item['Agency']).replace('.0', '').strip()
                    mapping_dict[(rt_key, ag_key)] = str(r_item['DRCODE']).strip()

            # 2. Read raw dataframe via pandas for smart Route & Agency detection
            df_input_raw = pd.read_excel(uploaded_file, header=None)

            # Find FG Row & Col dynamically
            fg_row, fg_col = -1, -1
            for r in range(df_input_raw.shape[0]):
                for c in range(df_input_raw.shape[1]):
                    val = str(df_input_raw.iloc[r, c]).strip().upper()
                    if "FG" in val:
                        fg_row, fg_col = r, c
                        break
                if fg_row != -1:
                    break

            if fg_row == -1:
                st.error("❌ Input file mein 'FG' header nahi mila.")
            else:
                # --- ROUTE DETECTION LOGIC ---
                route_num = "22"  # Default fallback
                match_route = re.search(r'Route\s*\(?(\d+)\)?', uploaded_file.name, re.IGNORECASE)
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

                # --- AGENCY COLUMN DETECTION LOGIC ---
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

                # --- OPENPYXL WORKFLOW (FORMAT, STYLING & FORMULAS PRESERVED) ---
                wb = openpyxl.load_workbook(uploaded_file)
                ws = wb.active

                excel_fg_row = fg_row + 1  # openpyxl is 1-indexed
                excel_fg_col = fg_col + 1
                
                # Check if DRCODE column already exists
                existing_drcode_col = None
                for col_idx in range(1, ws.max_column + 1):
                    cell_val = str(ws.cell(row=excel_fg_row, column=col_idx).value).strip().upper()
                    if cell_val == "DRCODE":
                        existing_drcode_col = col_idx
                        break

                # Insert DRCODE column right before FGCODE if not present
                target_col_idx = excel_fg_col
                if existing_drcode_col:
                    target_col_idx = existing_drcode_col
                else:
                    ws.insert_cols(excel_fg_col)
                    target_col_idx = excel_fg_col
                    ws.cell(row=excel_fg_row, column=target_col_idx, value="DRCODE")

                # Populate mapped DRCODE row by row using detected Route and Agency
                for row_idx in range(excel_fg_row + 1, ws.max_row + 1):
                    # agency_col is 0-indexed in pandas, so agency_col + 1 in openpyxl
                    raw_agency = ws.cell(row=row_idx, column=agency_col + 1).value
                    agency_str = str(raw_agency).replace('.0', '').strip() if raw_agency is not None else ""
                    
                    if agency_str and agency_str != "None":
                        lookup_key = (str(route_num), agency_str)
                        assigned_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_str}")
                        ws.cell(row=row_idx, column=target_col_idx, value=assigned_dr)

                st.success(f"✅ Route ({route_num}) & Agency detected successfully! DRCODE mapped and inserted right before FGCODE while keeping original file formatting intact.")

                # Save to buffer for download
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.download_button(
                    label="📥 Download Formatted & Mapped Excel File",
                    data=output_buffer.getvalue(),
                    file_name="Formatted_Mapped_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
