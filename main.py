from logging import exception

import pandas as pd
from fpdf import FPDF
from pathlib import Path
import streamlit as st
import os



UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "PDF"


def excel_to_PDF(filepath):

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf.set_font(family="Times", style='B', size=20)
    pdf.cell(w=0, h=12, txt=f"Invoice no:{invoice_nr}", align="L", ln=1)
    pdf.set_font(family="Times", style='B', size=20)
    pdf.cell(w=0, h=12, txt=f"Date:{date}", align="L", ln=1)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=35, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=f"The total amount is {total_sum}", ln=1)
    pdf.set_font(family="Times", style="B", size=10)


    output_path = os.path.join(OUTPUT_FOLDER,f"{filename}.pdf")
    pdf.output(output_path)
    return output_path


st.title("Excel to PDF Converter")
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    file_path = os.path.join(UPLOAD_FOLDER, uploaded_file.name)
    with open(file_path, 'wb') as file:
        file.write(uploaded_file.getbuffer())
    st.success("File uploaded successfully")

if st.button("Conver to PDF"):
    try:
        pdf_path = excel_to_PDF(file_path)
        st.success("Conversion successful!")

        with open(pdf_path, "rb") as file:
            st.download_button(label="Download PDF", data=file, file_name=os.path.basename(pdf_path), mime="application/pdf")

    except Exception as e:
        st.error(f"Error: {e}")





