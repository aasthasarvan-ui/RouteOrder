import streamlit as st
import pandas as pd
import os
import re

st.set_page_config(page_title="Auto DRCODE Mapping Tool", layout="wide")
st.title("💼 Smart Input & Master DRCODE Mapping Hub")

# 1. File Uploaders for Streamlit UI
st.subheader("Upload Files")
uploaded_file = st.file_uploader("Upload Input Demand Excel File", type=["xlsx", "xls"])
master_file_input = st.file_uploader("Upload Master Route File (Optional - agar GitHub mein nahi hai)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 2. Load Master Route File
        master_df = None
        master_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"
        
        if master_file_input is not None:
            master_df = pd.read_excel(master_file_input, sheet_name="Master Data", header=2)
            st.success("Master file uploaded successfully from device!")
        elif os.path.exists(master_path):
            master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
            st.info("Master file loaded automatically from repository.")
        else:
            st.error("⚠️ Master file nahi mili! Kripya Master file upload karein ya repository mein rakhein.")
        
        if master_df is not None:
            # Clean Master columns
            master_df.columns = master_df.columns.astype(str).str.strip()
            
            # Build Master Mapping Dictionary: (Route, Agency) -> DRCODE
            mapping_dict = {}
            if 'Route' in master_df.columns and 'Agency' in master_df.columns and 'DRCODE' in master_df.columns:
                master_subset = master_df[['Route', 'Agency', 'DRCODE']].dropna(subset=['DRCODE'])
                for _, r_item in master_subset.iterrows():
                    rt_key = str(r_item['Route']).replace('.0', '').strip()
                    ag_key = str(r_item['Agency']).replace('.0', '').strip()
                    mapping_dict[(rt_key, ag_key)] = str(r_item['DRCODE']).strip()

            # 3. Read Input File (Raw for detection, standard for manipulation)
            df_input_raw = pd.read_excel(uploaded_file, header=None)
            df_input = pd.read_excel(uploaded_file)
            df_input.columns = df_input.columns.astype(str).str.strip()

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
                # Detect Route Number from filename or header rows
                route_num = "22" # Default fallback
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
                            if cell_val != "" and 1 <= len(cell_val) <= 3 and any(char.isdigit() for char in cell_val):
                                route_num = cell_val
                                break
                        if route_num != "22":
                            break

                # Detect Agency Column dynamically
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

                # 4. Map DRCODE row-by-row based on Route and Agency
                drcodes_list = []
                for idx, row in df_input.iterrows():
                    raw_agency = df_input_raw.iloc[fg_row + 1 + idx, agency_col] if (fg_row + 1 + idx) < df_input_raw.shape[0] else None
                    agency_str = str(raw_agency).replace('.0', '').strip() if pd.notna(raw_agency) else ""
                    
                    lookup_key = (str(route_num), agency_str)
                    assigned_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_str}" if agency_str else "UNKNOWN")
                    drcodes_list.append(assigned_dr)

                # 5. Insert / Position DRCODE right before FGCODE column
                if 'DRCODE' in df_input.columns:
                    df_input = df_input.drop(columns=['DRCODE'])
                df_input['DRCODE'] = drcodes_list

                drc_col = df_input.pop('DRCODE')
                
                # Find FGCODE / Material code column name dynamically
                found_fg_col = None
                for col in df_input.columns:
                    if "FG" in str(col).upper() or "MATERIAL" in str(col).upper():
                        found_fg_col = col
                        break

                if found_fg_col:
                    fg_idx = df_input.columns.get_loc(found_fg_col)
                    df_input.insert(fg_idx, 'DRCODE', drc_col)
                    st.success(f"✅ Route ({route_num}) detected. DRCODE successfully mapped and inserted right before '{found_fg_col}'!")
                else:
                    df_input['DRCODE'] = drc_col
                    st.success(f"✅ Route ({route_num}) detected. DRCODE mapped and added to file.")

                # 6. Preview & Download Output
                st.subheader("Preview Processed Data (First 10 Rows)")
                st.dataframe(df_input.head(10))

                output_file_name = "Processed_Mapped_Output.xlsx"
                df_input.to_excel(output_file_name, index=False)
                
                with open(output_file_name, "rb") as f:
                    st.download_button(
                        label="📥 Download Mapped Excel File",
                        data=f,
                        file_name="Mapped_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
