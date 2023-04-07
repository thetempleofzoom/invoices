import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#note fp syntax (no slash required before folder)
filenames = glob.glob("files/*.xlsx")



for file in filenames:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    contents = pd.read_excel(file, sheet_name="Sheet 1")
    path = Path(file).stem
    invoice_number = path[:5]    #alternate -> file.split('\\')[1][:5]
    date = path[6:]   #alternate -> file.split('-')[1][:9]
    
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=f'Invoice Number: {invoice_number}', align="L", ln=1)
    pdf.cell(w=0, h=12, txt=f'Date: {date}', align="L", ln=1)
    pdf.output(f'PDF/{path}.pdf')