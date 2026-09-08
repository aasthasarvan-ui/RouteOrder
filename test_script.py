import streamlit as st
import pandas as pd
import openpyxl
import os
import re
import io

st.set_page_config(page_title="Auto DRCODE Mapping Tool", layout="wide")
st.title("💼 Smart Input & Master DRCODE Mapping Hub (Format Preserved)")

master_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"

# --- SIDEBAR: MANAGE MASTER DATA (ROBUST ADD & DELETE) ---
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
                    r_clean = str(new_route).replace('.0', '').strip()
                    a_clean = str(new_agency).replace('.0', '').strip()
                    ag2_clean = str(new_agency2).strip() if new_agency2 else a_clean
                    bp_clean = str(new_bp).strip()
                    dr_clean = str(new_drcode).strip()
                    
                    wb_m = openpyxl.load_workbook(master_path)
                    if "Master Data" in wb_m.sheetnames:
                        ws_m = wb_m["Master Data"]
                        
                        # Check duplicate and find last row accurately
                        is_duplicate = False
                        max_used_row = 2
                        
                        for r in range(3, ws_m.max_row + 1):
                            cell_r = ws_m.cell(row=r, column=1).value
                            cell_a = ws_m.cell(row=r, column=2).value
                            
                            if cell_r is not None or cell_a is not None:
                                max_used_row = r
                                
                            if cell_r is not None and cell_a is not None:
                                match_r = str(cell_r).replace('.0', '').strip()
                                match_a = str(cell_a).replace('.0', '').strip()
                                if match_r == r_clean and match_a == a_clean:
                                    is_duplicate = True
                                    break
                                    
                        if is_duplicate:
                            st.sidebar.error(f"⚠️ Duplicate Error: Route '{new_route}' aur Agency '{new_agency}' pehle se Master File mein maujood hain!")
                        else:
                            # Insert right after the last active row to keep formatting intact
                            target_row = max_used_row + 1
                            ws_m.cell(row=target_row, column=1, value=r_clean)
                            ws_m.cell(row=target_row, column=2, value=a_clean)
                            ws_m.cell(row=target_row, column=3, value=bp_clean)
                            ws_m.cell(row=target_row, column=4, value=ag2_clean)
                            ws_m.cell(row=target_row, column=5, value=dr_clean)
                            
                            wb_m.save(master_path)
                            st.sidebar.success("🎉 Naya record successfully Master File mein add ho gaya!")
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
                        st.sidebar.success(f"🗑️ Record successfully delete ho gaya (Route: {sel_route}, Agency: {sel_agency})!")
                        st.rerun()
                    else:
                        st.sidebar.error("❌ Record delete karne mein match nahi mila.")
            else:
                st.sidebar.error("❌ Master file ke columns read nahi ho paaye.")
        except Exception as e:
            st.sidebar.error(f"❌ Error loading master file: {e}")
    else:
        st.sidebar.warning("⚠️ Master file project folder mein nahi mili.")

# --- DOWNLOAD UPDATED MASTER FILE BUTTON ---
if os.path.exists(master_path):
    with open(master_path, "rb") as master_f:
        master_bytes = master_f.read()
        st.sidebar.download_button(
            label="📥 Download Updated Master File",
            data=master_bytes,
            file_name="Business_Partners_Master_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")

# --- MAIN APP SECTION: INPUT DEMAND FILE PROCESSING ---
uploaded_file = st.file_uploader("Upload Input Demand Excel File", type=["xlsx", "xls"])
master_file_input = st.file_uploader("Upload Master Route File (Optional - agar project folder mein nahi hai)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        master_df = None
        mapping_dict = {}
        
        if master_file_input is not None:
            try:
                master_df = pd.read_excel(master_file_input, sheet_name="Master Data", header=2)
                st.success("Master file uploaded successfully from device!")
            except Exception as e:
                st.warning(f"⚠️ Uploaded master file read karne mein issue aaya: {e}")
                
        elif os.path.exists(master_path):
            try:
                master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
                st.info("Master file loaded automatically from project folder!")
            except Exception as e:
                st.warning(f"⚠️ Project folder ki master file read karne mein issue aaya: {e}")
        else:
            st.warning("⚠️ Master file nahi mili. Sabhi agencies ke liye default 'NEW_CUST_' code assign kiya jayega.")

        if master_df is not None:
            master_df.columns = master_df.columns.astype(str).str.strip()
            if 'Route' in master_df.columns and 'Agency' in master_df.columns and 'DRCODE' in master_df.columns:
                master_subset = master_df[['Route', 'Agency', 'DRCODE']].dropna(subset=['DRCODE'])
                for _, r_item in master_subset.iterrows():
                    rt_key = str(r_item['Route']).replace('.0', '').strip()
                    ag_key = str(r_item['Agency']).replace('.0', '').strip()
                    mapping_dict[(rt_key, ag_key)] = str(r_item['DRCODE']).strip()

        df_input_raw = pd.read_excel(uploaded_file, header=None)

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
            route_num = "22"
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

            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb.active

            excel_fg_row = fg_row + 1
            excel_fg_col = fg_col + 1
            
            existing_drcode_col = None
            for col_idx in range(1, ws.max_column + 1):
                cell_val = str(ws.cell(row=excel_fg_row, column=col_idx).value).strip().upper()
                if cell_val == "DRCODE":
                    existing_drcode_col = col_idx
                    break

            target_col_idx = excel_fg_col
            if existing_drcode_col:
                target_col_idx = existing_drcode_col
            else:
                ws.insert_cols(excel_fg_col)
                target_col_idx = excel_fg_col
                ws.cell(row=excel_fg_row, column=target_col_idx, value="DRCODE")

            for row_idx in range(excel_fg_row + 1, ws.max_row + 1):
                raw_agency = ws.cell(row=row_idx, column=agency_col + 1).value
                agency_str = str(raw_agency).replace('.0', '').strip() if raw_agency is not None else ""
                
                if agency_str and agency_str != "None":
                    lookup_key = (str(route_num), agency_str)
                    assigned_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_str}")
                    ws.cell(row=row_idx, column=target_col_idx, value=assigned_dr)

            st.success(f"✅ Route ({route_num}) & Agency detected successfully! DRCODE mapped and inserted right before FGCODE while keeping original file formatting intact.")

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
