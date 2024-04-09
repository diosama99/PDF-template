from fpdf import FPDF
import pandas as pd

pdf = FPDF(orientation="P", unit="mm", format="A4")
pdf.set_auto_page_break(auto=False, margin=0)

df = pd.read_csv("003 topics.csv")


for index, row in df.iterrows():

    pdf.add_page()

    # Set the header
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=row["Topic"], align='L', ln=1)
    pdf.line(10, 20, 200, 20)

    #Set lines
    y = 20
    for i in range(38):
        pdf.line(10, y, 200, y)
        y += 7

    #Set the footer
    pdf.ln(260)
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(180, 0, 180)
    pdf.cell(w=0, h=10, txt=row["Topic"], align='R')

    for i in range(row["Pages"]-1):
        pdf.add_page()

        # Set the footer
        pdf.ln(272)
        pdf.set_font(family="Times", style="B", size=14)
        pdf.set_text_color(180, 0, 180)
        pdf.cell(w=0, h=10, txt=row["Topic"], align='R')

        # Set lines
        y = 20
        for i in range(38):
            pdf.line(10, y, 200, y)
            y += 7

pdf.output("output.pdf")