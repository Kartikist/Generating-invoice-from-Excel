import pandas as pd
import glob as g
from fpdf import FPDF
from pathlib import Path

filepaths = g.glob('invoices/*.xlsx')


for filepath in filepaths:
    pdf = FPDF(orientation="portrait", unit="mm",format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no, date = filename.split('-')
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice No: {invoice_no}")
    pdf.ln(8)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=4)
    
    df = pd.read_excel(filepath ,sheet_name= 'Sheet 1')
    
    # header
    columns = df.columns
    columns = [i.replace("_"," ").title() for i in columns]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=8, txt=columns[0], border=1 )
    pdf.cell(w=70, h=8, txt=columns[1], border=1 )
    pdf.cell(w=35, h=8, txt=columns[2], border=1 )
    pdf.cell(w=30, h=8, txt=columns[3], border=1 )
    pdf.cell(w=30, h=8, txt=columns[4], border=1,ln=1 )
    
    # table rows
    for i, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1, )
        pdf.cell(w=70, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]),border=1, ln=1)
    
    total_sum = df["total_price"].sum()
    pdf.set_font(family='Times', size=10)
    pdf.cell(w=70, h=8, txt="",border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="",border=1)
    pdf.cell(w=30, h=8, txt="",border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum),border=1, ln=1)
     
    pdf.set_font(family='Times', size=10, style='B')    
    pdf.cell(w=30, h=8, txt=f"Total amount to be paid: {total_sum}", ln=50)
    pdf.set_font(family='Times', size=30, style='BIU') 
    pdf.cell(w=30, h=20, txt=f"The PDF company")
    
    pdf.output(f'pdfs/{filename}.pdf')