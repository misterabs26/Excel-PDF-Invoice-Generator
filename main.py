import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


excel_files = glob.glob("invoices/*.xlsx")

for excel_file in excel_files:
    excel_file = Path(excel_file)
    df = pd.read_excel(excel_file,sheet_name="Sheet 1")

    filename = excel_file.stem
    invoice_no, invoice_date = filename.split("-")

    pdf = FPDF(orientation="P",unit="mm",format="letter")
    pdf.add_page()
    pdf.set_font(family="Times", style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no.: {invoice_no}", ln=1)
    pdf.cell(w=50, h=8,txt=f"Date: {invoice_date}", ln=1)


    pdf.output(f"pdf/{filename}.pdf")