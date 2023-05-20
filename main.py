import pandas as pd
import openpyxl
from fpdf import FPDF
import glob
from pathlib import Path

# Get everything with an .xlsx extension
filepaths = glob.glob("invoices/*.xlsx")

# Read evey Excel file in the filepaths
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="L", unit="mm", format="A4")

    pdf.add_page()

    file_names = Path(filepath).stem.split("-")
    invoice_number = file_names[0]
    date = file_names[1]

    pdf.set_font("Times", style="B", size=16)
    pdf.cell(h=8, w=50, align="L", txt=f"Invoice nr. {invoice_number}", ln=1)
    pdf.cell(h=8, w=50, align="L", txt=f"Date: {date}", ln=1)

    pdf.output(f"PDFs/Invoice {invoice_number}.pdf")
