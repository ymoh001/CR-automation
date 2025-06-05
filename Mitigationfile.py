import os
import time
import pandas as pd
from datetime import datetime

start_time = time.time()

exchange_file = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Exchange rates.xlsx"
exchange_df = pd.read_excel(exchange_file)
exchange_df.columns = [str(c).strip() for c in exchange_df.columns]
# Only use rows ending in '/MYR'
exchange_df['Currency.1'] = exchange_df['Currency.1'].astype(str).str.strip()
exchange_df = exchange_df[exchange_df['Currency.1'].str.upper().str.endswith('/MYR')]
exchange_df['Base Currency'] = exchange_df['Currency.1'].str.split('/').str[0].str.strip().str.upper()
# Remove non-date columns
numeric_cols = [col for col in exchange_df.columns if col not in ['Currency', 'Currency.1', 'Base Currency'] and exchange_df[col].dtype != 'O']
if not numeric_cols:
    # Fallback if all are object: just skip the first two
    numeric_cols = [col for col in exchange_df.columns if col not in ['Currency', 'Currency.1', 'Base Currency']]
selected_month = numeric_cols[-2]  # use the rightmost
print("Using selected_month:", selected_month)

exchange_dict = {
    row['Base Currency']: float(row[selected_month])
    for idx, row in exchange_df.iterrows()
    if not pd.isnull(row[selected_month])
}

main_folder = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Sample file 2"
output_folder = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Python outputs"
output_file = os.path.join(output_folder, "mitigation.xlsx")
part_headers = [
    "Data entry", "CR", "Model", "Project", "LTB", "CM", "Dyson PN", "DESCRIPTION", "Supplier", "Commodity", "Currency", "Unit Price", 
    "OPO", "OPO $", "SOH", "SOH $", "Other mitigation cost $", "Total Exposure Qty", "Total Exposure $", "Total exposure in MYR", "RDD Remark", "Other Remark"
]
# mapping from target header to part table header (if any)
header_map = {
    "Dyson PN": "Dyson PN no rev",
    "DESCRIPTION": "DESCRIPTION",
    "Supplier": "Supplier",
    "Commodity": "Commodity",
    "Currency": "Currency",
    "Unit Price": "U/Price",
    "OPO": "Balance OPO(Qty)",
    "OPO $": "Balance OPO ($)",
    "SOH": "Balance SOH (Qty)",
    "SOH $": "Balance SOH ($)",
    "Other mitigation cost $": "Other mitigation cost ($)",
    "Total Exposure Qty": "Total Mitigation in Progress (Qty)",
    "Total Exposure $": "Total Mitigation in Progress ($)",
    "LTB": "LT (Wks)",
}

def extract_metadata(df):
    # Scan first 20 rows for the relevant keywords in column A, return value in column C
    meta = {"CR": "", "Model": "", "Project": "", "CM": ""}
    for i in range(20):
        keycell = str(df.iloc[i, 0]).strip().lower() if pd.notnull(df.iloc[i, 0]) else ""
        valcell = str(df.iloc[i, 2]).strip() if pd.notnull(df.iloc[i, 2]) else ""
        if "contract manufacturing" in keycell:
            meta["CM"] = valcell
        elif "model name" in keycell:   # e.g. "SV25 X285"
            words = valcell.split()
            if len(words) > 0:
                meta["Model"] = words[0]
            if len(words) > 1:
                meta["Project"] = words[1]
        elif "remarks" in keycell:
            meta["CR"] = valcell
    return meta

def compute_myr(row):
    currency = str(row["Currency"]).strip().upper()
    exposure = pd.to_numeric(row["Total Exposure $"], errors="coerce")
    rate = exchange_dict.get(currency, 1.0)
    print(f"Row Currency: {currency} | Exposure: {exposure} | Rate: {rate}")
    if pd.isnull(exposure):
        return ""
    return round(exposure * rate, 2)

all_rows = []

for filename in os.listdir(main_folder):
    if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
        continue
    filepath = os.path.join(main_folder, filename)
    print(f"Processing: {filename}")
    try:
        df = pd.read_excel(filepath, sheet_name="Appendix 2", header=None)
        metadata = extract_metadata(df)
        today_str = datetime.now().strftime("%Y-%m-%d")

        # Find and extract header rows first
        # ----------------------
        # Block 1
        header_row1 = 19
        block1 = df.iloc[header_row1:300, 0:15].copy()
        block1.columns = block1.iloc[0]   # Header
        block1 = block1.iloc[1:].reset_index(drop=True)
        # Block 2
        header_row2 = 19
        block2 = df.iloc[header_row2:300, 117:128].copy()
        block2.columns = block2.iloc[0]   # Header
        block2 = block2.iloc[1:].reset_index(drop=True)
        # NB: adjust header_row1/header_row2 if your actual header is not 19!

        length = min(len(block1), len(block2))
        for i in range(length):
            row1 = block1.iloc[i]
            row2 = block2.iloc[i]
            output_row = {
                "Data entry": today_str,
                "CR": metadata.get("CR", ""),
                "Model": metadata.get("Model", ""),
                "Project": metadata.get("Project", ""),
                "CM": metadata.get("CM", "")
            }
            # Fill part_headers from both blocks using your header_map
            for h in part_headers:
                if h in output_row:
                    continue  # already filled from metadata
                val = ""
                # First try from block1
                col1 = header_map.get(h)
                if col1 and col1 in row1.index:
                    val = row1[col1]
                # If not in block1, try block2
                if not val and col1 and col1 in row2.index:
                    val = row2[col1]
                output_row[h] = val
            all_rows.append(output_row)

    except Exception as e:
        print(f"Error processing {filename}: {e}")

if all_rows:
    consolidated_df = pd.DataFrame(all_rows, columns=part_headers)
    consolidated_df = consolidated_df[
    consolidated_df["Dyson PN"].notnull() & 
    (consolidated_df["Dyson PN"] != "") &
    (pd.to_numeric(consolidated_df["Total Exposure $"], errors='coerce') != 0) &
    (pd.to_numeric(consolidated_df["Total Exposure $"], errors='coerce').notnull())
    ]
    consolidated_df["Total exposure in MYR"] = consolidated_df.apply(compute_myr, axis=1)
    consolidated_df.to_excel(output_file, index=False)
    print(f"Saved consolidated parts list to {output_file}")
else:
    print("No data rows found to consolidate.")

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Elapsed time: {elapsed_time:.2f} seconds")
