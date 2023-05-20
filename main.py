import pandas as pd
import openpyxl
from fpdf import FPDF as pdf
import glob

# Get everything with an .xlsx extension
filepaths = glob.glob("invoices/*.xlsx")

# Read evey Excel file in the filepaths
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
