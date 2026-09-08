import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Sales Order Auto DRCODE Mapping", layout="wide")
st.title("Sales Order Automation with Auto DRCODE Mapping")

# --- Step 1: File Uploaders ---
st.subheader("1. File Uploads")
uploaded_file = st.file_uploader("Upload Input Demand Excel File", type=["xlsx", "xls"])
master_file_input = st.file_uploader("Upload Master Route File (Optional agar GitHub me hai)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        master_df = None
        master_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"
        
        # Master file load karne ka logic (header=2 zaroori hai kyunki actual columns row 3 par hain)
        if master_file_input is not None:
            master_df = pd.read_excel(master_file_input, sheet_name="Master Data", header=2)
            st.success("Master file uploaded successfully from device!")
        elif os.path.exists(master_path):
            master_df = pd.read_excel(master_path, sheet_name="Master Data", header=2)
            st.info("Master file loaded automatically from project repository.")
        else:
            st.error("⚠️ Master file nahi mili! Kripya ya toh GitHub par file rakhein ya upar upload option se select karein.")
        
        if master_df is not None:
            # --- Step 2: Read Input File ---
            df_input = pd.read_excel(uploaded_file)
            
            # Clean column names (remove extra spaces)
            df_input.columns = df_input.columns.astype(str).str.strip()
            master_df.columns = master_df.columns.astype(str).str.strip()
            
            # --- Step 3: Matching Keys ---
            input_route_col = 'Route'     
            input_agency_col = 'Agency'   
            
            master_route_col = 'Route'
            master_agency_col = 'Agency'
            
            # Prepare Master for Merge
            df_temp_master = master_df[[master_route_col, master_agency_col, 'DRCODE']].drop_duplicates()
            
            # Agar purana DRCODE column hai toh hata do
            if 'DRCODE' in df_input.columns:
                df_input = df_input.drop(columns=['DRCODE'])
                
            # Merge (VLOOKUP equivalent based on Route + Agency)
            df_merged = pd.merge(
                df_input,
                df_temp_master,
                left_on=[input_route_col, input_agency_col],
                right_on=[master_route_col, master_agency_col],
                how='left'
            )
            
            # Cleanup extra columns if names differed
            if master_route_col != input_route_col and master_route_col in df_merged.columns:
                df_merged = df_merged.drop(columns=[master_route_col])
            if master_agency_col != input_agency_col and master_agency_col in df_merged.columns:
                df_merged = df_merged.drop(columns=[master_agency_col])
                
            df_input = df_merged
            
            # --- Step 4: Positioning Logic (Insert DRCODE right before FGCODE) ---
            fg_code_column_name = 'FGCODE' # Agar material code ka naam kuch aur ho toh yahan change kar sakte hain
            
            if 'DRCODE' in df_input.columns:
                drc_col = df_input.pop('DRCODE')
                if fg_code_column_name in df_input.columns:
                    fg_idx = df_input.columns.get_loc(fg_code_column_name)
                    df_input.insert(fg_idx, 'DRCODE', drc_col)
                else:
                    df_input['DRCODE'] = drc_col
            
            st.success("✅ DR Codes successfully mapped and inserted right before FGCODE!")
            
            # --- Step 5: Preview Result ---
            st.subheader("Processed Data Preview (First 10 Rows)")
            st.dataframe(df_input.head(10))
            
            # Download Button for Final Processed File
            output_file_name = "Processed_Demand_Output.xlsx"
            df_input.to_excel(output_file_name, index=False)
            
            with open(output_file_name, "rb") as f:
                st.download_button(
                    label="📥 Download Processed Excel File",
                    data=f,
                    file_name="Processed_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"❌ Error during processing: {e}")
