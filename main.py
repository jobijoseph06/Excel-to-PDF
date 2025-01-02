import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", style='B', size=20)
    pdf.set_text_color(100,100,100)
    pdf.cell(w=0, h=12, txt=f"Invoice_nr:{invoice_nr}", align="L", ln=1)
    pdf.output(f"PDF/{filename}.pdf")