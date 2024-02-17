import pandas as pd
"""
glob function is used to search for files that match a specific file pattern or name 
"""
import glob


filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
