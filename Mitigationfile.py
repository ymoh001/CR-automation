import os
import pandas as pd

main_folder = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Sample file 2"
output_file = os.path.join(main_folder, "all_parts_consolidated.xlsx")

part_headers = [
    "Data entry", "CR", "Model", "Project", "LTB", "CM", "Dyson PN", "DESCRIPTION", "Supplier", "Commodity", "Currency", "Unit Price", 
    "OPO", "OPO $", "SOH", "SOH $", "Other mitigation cost $", "Total Exposure Qty", "Total Exposure $", "Total exposure in MYR", "RDD Remark", "Other Remark"
]

# Helper: Try to match your target header to one in the data by substring
def find_best_match(target, cols):
    for c in cols:
        if target.lower().replace(" ", "") in c.lower().replace(" ", ""):
            return c
    # If not found, return None
    return None

all_rows = []

for filename in os.listdir(main_folder):
    if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
        continue
    filepath = os.path.join(main_folder, filename)
    print(f"Processing: {filename}")
    try:
        # Read the relevant sheet without headers (we'll treat actual headers as a row in the block)
        df = pd.read_excel(filepath, sheet_name="Appendix 2", header=None)
        # 1. First "table": A20:O300 (pandas: rows 19-299, cols 0-14)
        df1 = df.iloc[19:300, 0:15].copy()
        # 2. Second table: DN20:DT300
        df2 = df.iloc[19:300, 117:124].copy()
        # Get the header row for each
        df1.columns = df1.iloc[0]
        df1 = df1.iloc[1:]
        df2.columns = df2.iloc[0]
        df2 = df2.iloc[1:]

        # for each, try to map to your desired output headers.
        for working_df in [df1, df2]:
            if working_df.empty:
                continue
            mapped = pd.DataFrame()
            for h in part_headers:
                match = find_best_match(h, working_df.columns)
                if match:
                    mapped[h] = working_df[match]
                else:
                    mapped[h] = ""     # Or pd.NA, or None, as you prefer
            # Optionally, add a column for the filename (file source)
            mapped["Source file"] = filename
            all_rows.append(mapped)
    except Exception as e:
        print(f"Error processing {filename}: {e}")

if all_rows:
    consolidated_df = pd.concat(all_rows, ignore_index=True)
    consolidated_df.to_excel(output_file, index=False)
    print(f"Saved consolidated parts list to {output_file}")
else:
    print("No data rows found to consolidate.")