import os
import re
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Start the timer
start_time = time.time()

# Define the main folder path
main_folder = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Sample file 2"

# Define the output file path
output_file = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Consolidate Exposure\Test setup 2\Masterfile May '25.xlsx"

# Headers for the master file
headers = [
    'Initial Claim submission Date','CR Number','CR Description','EOP Strategy','CM','EOP Declaration Timing','Last Time Build','Dyson PIC','Product Category','Project','Model',	
    'Initial Submission','Claim Received (RM)','Claim Accepted (RM)','Claim value pending SAF/PR approval (RM)','Claim Avoided (RM)','Claim in Progress (RM)','WIP (RM/USD)',	
    'Remark/Current Status','One Time Settlement','Claim Status','Finance Status','CM Claim No (Commercial Title)','PR Number','PO Number','GR Status','GR Amount','Accrued/GR Amt','Provision','Check'
]
