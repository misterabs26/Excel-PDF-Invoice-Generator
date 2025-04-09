import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

excel_files = glob.glob("invoices/*.xlsx")

def table_header(data):
    pdf.set_font("Times", style="B", size=13)
    table_width = pdf.w - 20
    no_of_col = len(data.columns)
    col_width = table_width/no_of_col

    columns = [col.replace("_"," ") for col in data.columns]

    for header in columns:
        header_name = header.title()
        pdf.cell(col_width,10,header_name,border=1,align="C")

    pdf.ln(1)


def populate_table(y_pos,df_row):
    pdf.set_font("Times", style="", size=10)
    table_width = pdf.w - 20
    no_of_col = len(df_row)
    col_width = table_width/no_of_col

    pdf.set_y(y_pos)
    pdf.cell(col_width,10,str(df_row["product_id"]),border=1)
    pdf.cell(col_width, 10, str(df_row["product_name"]), border=1)
    pdf.cell(col_width, 10, str(df_row["amount_purchased"]), border=1)
    pdf.cell(col_width, 10, str(df_row["price_per_unit"]), border=1)
    pdf.cell(col_width, 10, str(df_row["total_price"]), border=1,ln=1)


for excel_file in excel_files:
    pdf = FPDF(orientation="L", unit="mm", format="letter")
    excel_file = Path(excel_file)

    filename = excel_file.stem
    invoice_no, invoice_date = filename.split("-")

    pdf.add_page()
    pdf.set_font(family="Times", style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no.: {invoice_no}", ln=1)
    pdf.cell(w=50, h=8,txt=f"Date: {invoice_date}", ln=1)
    pdf.ln(5)

    df = pd.read_excel(excel_file, sheet_name="Sheet 1")
    # Table headers
    table_header(df)

    # Table contents/data
    header_pos = pdf.get_y()
    pdf.set_y(header_pos+9)

    for i, row in df.iterrows():
        populate_table(pdf.get_y(),row)


    pdf.output(f"pdf/{filename}.pdf")