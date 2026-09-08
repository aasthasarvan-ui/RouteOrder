import pandas as pd
import os
import re

def test_mapping_logic():
    print("--- [TEST] Starting Smart DR Code Mapping Test ---")
    
    # 1. File names
    input_file = "rt222222 (003) with new row.xlsx"
    master_file = "Business_Partners_Master_Original_Keys_Restored.xlsx"
    
    if not os.path.exists(input_file) or not os.path.exists(master_file):
        print("[Error] Input file ya Master file nahi mili! Please check file names.")
        return
    
    try:
        # 2. Read Input File without header to locate FG and structure dynamically
        df_input_raw = pd.read_excel(input_file, header=None)
        
        # Find FG Row & Col
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
            print("[Error] 'FG' header not found in input file.")
            return

        print(f"[Info] FG found at Row {fg_row}, Col {fg_col}")

        # 3. Detect Route Number (from filename or header rows)
        route_num = "22" # Default fallback
        match_route = re.search(r'Route\s*\(?(\d+)\)?', input_file, re.IGNORECASE)
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

        print(f"[Info] Detected Route No: {route_num}")

        # 4. Detect Agency Column
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

        print(f"[Info] Detected Agency Column Index: {agency_col}")

        # 5. Load Master File (header=2 as observed in structure)
        df_master = pd.read_excel(master_file, sheet_name="Master Data", header=2)
        df_master.columns = df_master.columns.astype(str).str.strip()

        # Extract rows and map Route + Agency -> DRCODE
        # Re-read input file standard dataframe for manipulation
        df_input = pd.read_excel(input_file)
        df_input.columns = df_input.columns.astype(str).str.strip()

        # Let's build a clean mapping dictionary from Master Data
        master_subset = df_master[['Route', 'Agency', 'DRCODE']].dropna(subset=['DRCODE'])
        master_subset['Route'] = master_subset['Route'].astype(str).str.replace('.0', '').str.strip()
        master_subset['Agency'] = master_subset['Agency'].astype(str).str.replace('.0', '').str.strip()
        
        mapping_dict = {}
        for _, row in master_subset.iterrows():
            key = (row['Route'], row['Agency'])
            mapping_dict[key] = str(row['DRCODE']).strip()

        # Assign DRCODE to input data row by row based on detected Route and Agency column
        drcodes_list = []
        # Find which column in df_input corresponds to agency_col
        # Or we can match using standard column name if present, else position
        for idx, row in df_input.iterrows():
            # Get agency value from the detected agency column position
            raw_agency = df_input_raw.iloc[fg_row + 1 + idx, agency_col] if (fg_row + 1 + idx) < df_input_raw.shape[0] else None
            agency_str = str(raw_agency).replace('.0', '').strip() if pd.notna(raw_agency) else ""
            
            lookup_key = (str(route_num), agency_str)
            assigned_dr = mapping_dict.get(lookup_key, f"NEW_CUST_{agency_str}" if agency_str else "UNKNOWN")
            drcodes_list.append(assigned_dr)

        # 6. Insert / Update DRCODE column right before FGCODE
        if 'DRCODE' in df_input.columns:
            df_input = df_input.drop(columns=['DRCODE'])
            
        df_input['DRCODE'] = drcodes_list

        fg_code_column_name = 'FGCODE' # Or your material code column
        drc_col = df_input.pop('DRCODE')
        
        # Find FG column in df_input
        found_fg_col = None
        for col in df_input.columns:
            if "FG" in str(col).upper() or "MATERIAL" in str(col).upper():
                found_fg_col = col
                break

        if found_fg_col:
            fg_idx = df_input.columns.get_loc(found_fg_col)
            df_input.insert(fg_idx, 'DRCODE', drc_col)
            print(f"[Success] 'DRCODE' successfully inserted right before '{found_fg_col}'!")
        else:
            df_input['DRCODE'] = drc_col
            print("[Warning] FG column not found by name, added 'DRCODE' at the end.")

        # 7. Save Test Output
        output_test_file = "test_output_result.xlsx"
        df_input.to_excel(output_test_file, index=False)
        print(f"[Success] Test completed successfully! Output saved as '{output_test_file}'.")

    except Exception as e:
        print(f"[Error] Test failed due to: {e}")

if __name__ == "__main__":
    test_mapping_logic()
