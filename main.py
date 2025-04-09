import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

excel_files = glob.glob("invoices/*.xlsx")

# Setting column sizes
def get_column_widths(df, page_width):
    # Assign relative "weights" to each column
    weights = {
        "product_id": 1,
        "product_name": 2,   # Wider
        "amount_purchased": 1,
        "price_per_unit": 1,
        "total_price": 1
    }
    total_weight = sum(weights.values())

    # Convert weights into actual width in mm
    col_widths = {
        col: (weights[col]/total_weight) * page_width for col in df.columns
    }
    return col_widths


def table_header(data):
    pdf.set_font("Times", style="B", size=13)
    table_width = pdf.w - 20
    col_widths = get_column_widths(data,table_width)

    for col in data.columns:
        header = col.replace("_"," ").title()
        pdf.cell(col_widths[col],10,header,border=1,align="C")

    pdf.ln()


def populate_table(df_row):
    pdf.set_font("Times", style="", size=10)
    table_width = pdf.w - 20
    col_widths = get_column_widths(df_row.to_frame().T, table_width)

    for col in df_row.index:
        pdf.cell(col_widths[col],10,str(df_row[col]),border=1)
    pdf.ln()


# Iterate each Excel file
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
    for i, row in df.iterrows():
        populate_table(row)


    pdf.output(f"pdf/{filename}.pdf")