import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    pdf = FPDF(orientation='P', unit="mm", format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")
    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr.split(" ")[1]}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=1)
    pdf.ln(7)

    pdf.set_font(family="Times", size=10, style='B')
    pdf.cell(w=30, h=8, txt="Product ID", border=1)
    pdf.cell(w=40, h=8, txt="Product Name", border=1)
    pdf.cell(w=30, h=8, txt="Amount", border=1)
    pdf.cell(w=30, h=8, txt="Price per Unit", border=1)
    pdf.cell(w=30, h=8, txt="Total Price", border=1, ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # lst = list(df.columns)
    # lst = [item.replace("_", " ").title() for item in lst]
    # for item in lst:
    #     pdf.set_font(family="Times", size=10, style='B')
    #     pdf.cell(w=30, h=8, txt=lst[0], border=1)
    #     | - | - |
    # pdf.ln(8)

    for index, row in df.iterrows():

        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    totals = df['total_price'].sum()
    pdf.set_font(family="Times", size=8)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=130, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(totals), border=1)

    pdf.output(f"PDFs/{filename}.pdf")
