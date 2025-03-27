import pandas as pd
import glob

excel_files = glob.glob("invoices/*.xlsx")

for excel_file in excel_files:
    df = pd.read_excel(excel_file,sheet_name="Sheet 1")
    print(df)