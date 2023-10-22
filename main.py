import pandas as pd
import glob as g
from fpdf import FPDF
from pathlib import Path

filepaths = g.glob('invoices/*.xlsx')


for filepath in filepaths:
    df = pd.read_excel(filepath ,sheet_name= 'Sheet 1')
    pdf = FPDF(orientation="portrait", unit="mm",format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no = filename.split('-')[0]
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice No: {invoice_no}")
    pdf.output(f'pdfs/{filename}.pdf')