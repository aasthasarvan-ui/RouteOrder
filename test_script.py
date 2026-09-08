import streamlit as st
import pandas as pd
import os

st.title("Sales Order Automation with Auto DRCODE Mapping")

# 1. File Uploader for Input File
uploaded_file = st.file_uploader("Upload Input Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Master file ka path (jo aapke project folder me rakhi hai)
        master_file_path = "Business_Partners_Master_Original_Keys_Restored.xlsx"
        
        if not os.path.exists(master_file_path):
            st.error(f"Master file '{master_file_path}' nahi mili! Please project folder me rakhein.")
        else:
            # 2. Read Uploaded Input File
            df_input = pd.read_excel(uploaded_file)
            df_master = pd.read_excel(master_file_path, sheet_name="Master Data")
            
            # Clean column names
            df_input.columns = df_input.columns.astype(str).str.strip()
            df_master.columns = df_master.columns.astype(str).str.strip()
            
            # 3. Matching Keys
            input_route_col = 'Route'     # Input file ka route column
            input_agency_col = 'Agency'   # Input file ka agency column
            
            master_route_col = 'Route'
            master_agency_col = 'Agency'
            
            # Prepare Master for Merge
            df_temp_master = df_master[[master_route_col, master_agency_col, 'DRCODE']].drop_duplicates()
            
            # Agar purana DRCODE column hai toh hata do
            if 'DRCODE' in df_input.columns:
                df_input = df_input.drop(columns=['DRCODE'])
                
            # Merge (VLOOKUP equivalent)
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
            
            # 4. Positioning Logic: Insert DRCODE right before FGCODE
            fg_code_column_name = 'FGCODE' # Apne actual material code column ka naam yahan rakhein
            
            if 'DRCODE' in df_input.columns:
                drc_col = df_input.pop('DRCODE')
                if fg_code_column_name in df_input.columns:
                    fg_idx = df_input.columns.get_loc(fg_code_column_name)
                    df_input.insert(fg_idx, 'DRCODE', drc_col)
                else:
                    df_input['DRCODE'] = drc_col
            
            st.success("DR Codes successfully mapped and inserted!")
            
            # 5. Preview updated data on Streamlit UI
            st.dataframe(df_input.head(10))
            
            # 6. (Optional) Yahan se aapka baaki ka main processing logic shuru ho sakta hai
            # df_output = your_main_calculation_function(df_input)
            
            # Download button for processed file
            # ...
            
    except Exception as e:
        st.error(f"Error during processing: {e}")
