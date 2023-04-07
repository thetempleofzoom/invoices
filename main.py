import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#note fp syntax (no slash required before folder)
filenames = glob.glob("files/*.xlsx")



for file in filenames:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(False, margin=0)
    
    #get invoice number and date
    path = Path(file).stem
    invoice_number = path[:5]    #alternate -> file.split('\\')[1][:5]
    date = path[6:]   #alternate -> file.split('-')[1][:9]
    #add to top of pdf page
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=f'Invoice Number: {invoice_number}', align="L", ln=1)
    pdf.cell(w=0, h=12, txt=f'Date: {date}', align="L", ln=1)

    #table - note that each row has to be physically placed in pdf
    #will swap qty and unit price around to illustrate the point
    contents = pd.read_excel(file, sheet_name="Sheet 1")
    headers = list(contents)
    # very cool technique to swap places in list -> headers[2:4] = reversed(headers[2:4])

    #mapping raw headers to pretty ones
    pretty = pd.read_excel('pretty.xlsx')
    prettyheaders = pretty.set_index('raw').to_dict()['pretty']

    #generate header row
    pdf.set_font('courier', 'B', 12)
    pdf.set_text_color(100, 80, 80)
    pdf.cell(w=30, h=10, txt=str(prettyheaders[headers[0]]), border=1)
    pdf.cell(w=70, h=10, txt=str(prettyheaders[headers[1]]), border=1)
    pdf.cell(w=30, h=10, txt=str(prettyheaders[headers[3]]), border=1)
    pdf.cell(w=30, h=10, txt=str(prettyheaders[headers[2]]), border=1)
    pdf.cell(w=30, h=10, txt=str(prettyheaders[headers[4]]), border=1, ln=1)

    for index, row in contents.iterrows():
        pdf.set_font('courier', '', 12)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=10, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=10, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['total_price']), border=1, ln=1)

    #adding totals
    total = sum(contents['total_price'])
    pdf.set_font('courier', 'B', 12)
    pdf.set_text_color(100, 80, 80)
    pdf.set_x(110)
    pdf.cell(w=30, h=10, txt='Total')
    pdf.set_x(170)
    pdf.cell(w=30, h=10, txt=str(total), ln=1)
    pdf.ln(20)

    pdf.set_font('courier', 'B', 16)
    pdf.set_text_color(100, 80, 80)
    pdf.cell(w=30, h=10, txt=f'Total amt due is ${total}', ln=1)
    pdf.image("7109.jpg",w=30, h=20)
    pdf.output(f'PDF/{path}.pdf')