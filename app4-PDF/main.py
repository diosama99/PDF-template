from fpdf import FPDF
import pandas as pd

pdf = FPDF(orientation="P", unit="mm", format="A4")

df = pd.read_csv("003 topics.csv")

for index, row in df.iterrows():

    pdf.add_page()
    pdf.set_font(family="Times", style="B")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=row["Topic"], align='L', ln=1)
    pdf.line(10, 20, 200, 20)

    for i in range(row["Pages"]-1):
        pdf.add_page()

pdf.output("output.pdf")