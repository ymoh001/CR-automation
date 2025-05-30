import os
import re
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Start the timer
start_time = time.time()

# === USER CONFIGURATION ===
main_folder = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Sample file 2"
summary_filename = "master_summary.xlsx"

# Your desired header order (customize as needed)
headers = [
    'Initial Claim submission Date','CR Number','CR Description','EOP Strategy','CM','EOP Declaration Timing','Last Time Build','Dyson PIC',
    'Product Category','Project','Model', 'Initial Submission','Claim Received (RM)','Claim Accepted (RM)','Claim value pending SAF/PR approval (RM)',
    'Claim Avoided (RM)','Claim in Progress (RM)','WIP (RM/USD)','Remark/Current Status','One Time Settlement','Claim Status','Finance Status',
    'CM Claim No (Commercial Title)','PR Number','PO Number','GR Status','GR Amount','Accrued/GR Amt','Provision','Check'
]

# --- Function to Parse Model/Project from model cell ---
def parse_model_project(val):
    if isinstance(val, str):
        parts = val.strip().split()
        if len(parts) == 2:
            return parts[0], parts[1]
    return None, None

# --- Metadata Extractor ---
def extract_exposure_metadata(df):
    mapping = {}
    for i in range(6, 16):  # Rows A7 to F16 (pandas index 6:15)
        # Columns: A-B (0,2), E-F (4,5)
        cells = [
            (str(df.iloc[i, 0]).strip().lower() if pd.notnull(df.iloc[i, 0]) else "", df.iloc[i, 2]),
            (str(df.iloc[i, 4]).strip().lower() if pd.notnull(df.iloc[i, 4]) else "", df.iloc[i, 5])
        ]
        for k, v in cells:
            if not k:
                continue
            if "eop declare date" in k:
                mapping["EOP Declaration Timing"] = v
            elif "initial submission date" in k:
                mapping["Initial Claim submission Date"] = v
            elif "ltb week" in k:
                mapping["Last Time Build"] = v
            elif "initial submission value" in k:
                mapping["Initial Submission"] = v
            elif "contract manufacturing" in k:
                mapping["CM"] = v
            elif "currency" in k:
                mapping["Currency"] = v
            elif "category" in k:
                mapping["Product Category"] = v
            elif "exchange rate" in k:
                mapping["Exchange rate to MYR"] = v
            elif "model name" in k:
                model, project = parse_model_project(v)
                mapping["Model"] = model
                mapping["Project"] = project
            elif "cm claim no" in k:
                mapping["CM Claim No (Commercial Title)"] = v
            elif "dyson pic" in k:
                mapping["Dyson PIC"] = v
            elif "claim status" in k:
                mapping["Claim Status"] = v
            elif "cm pic" in k:
                mapping["CM PIC"] = v
            elif "pr number" in k:
                mapping["PR Number"] = v
            elif "remarks" in k:
                mapping["CR Number"] = v
            elif "po number" in k:
                mapping["PO Number"] = v
            elif "cr description" in k:
                mapping["CR Description"] = v
            elif "gr amount" in k:
                mapping["GR Amount"] = v
            elif "eop stratergy" in k or "eop strategy" in k:
                mapping["EOP Strategy"] = v
            elif "ranging out" in k:
                mapping["Ranging Out"] = v
            elif "cr-" in str(v).lower():
                mapping['CR Number'] = v  # Optionally, auto-fill CR Number if it matches

    return mapping

# --- MAIN SCRIPT ---
output_rows = []
for filename in os.listdir(main_folder):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        filepath = os.path.join(main_folder, filename)
        try:
            df = pd.read_excel(filepath, sheet_name=0, header=None)  # Read first sheet, no header
            meta = extract_exposure_metadata(df)
            row = []
            for h in headers:
                # Fallback to "" or None for missing values
                val = meta.get(h, "")
                row.append(val)
            output_rows.append(row)
            print(f"Processed: {filename}")
        except Exception as e:
            print(f"Error processing {filename}: {e}")

# Build summary DataFrame and save
summary_df = pd.DataFrame(output_rows, columns=headers)
summary_df.to_excel(os.path.join(main_folder, summary_filename), index=False)
print(f"Summary saved as {summary_filename} in {main_folder}")

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Elapsed time: {elapsed_time:.2f} seconds")