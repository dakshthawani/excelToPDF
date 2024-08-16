import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no, invoice_date = filename.split("-")
    pdf.set_font("Times",style="B",size=24)
    pdf.cell(w=0,h=12,txt=f"Invoice no. {invoice_no}",border=0,ln=1,align="L")
    pdf.set_font("Times", style="B", size=12)
    pdf.cell(w=0, h=12, txt=f"Invoice date {invoice_date}", border=0, ln=1, align="L")
    # pdf.output(f"PDF/{filename}.pdf")
    columns = list(df.columns)
    columns = [items.replace("_", " ").title() for items in columns]
    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w=30, h=10, border=1, txt=columns[0])
    pdf.cell(w=50, h=10, border=1, txt=columns[1])
    pdf.cell(w=40, h=10, border=1, txt=columns[2])
    pdf.cell(w=30, h=10, border=1, txt=columns[3])
    pdf.cell(w=40, h=10, border=1, txt=columns[4], ln=1)
    for index, row in df.iterrows():
        pdf.set_font("Times", size=12)
        pdf.cell(w=30, h=10, border=1, txt=str(row['product_id']))
        pdf.cell(w=50, h=10, border=1, txt=str(row['product_name']))
        pdf.cell(w=40, h=10, border=1, txt=str(row['amount_purchased']))
        pdf.cell(w=30, h=10, border=1, txt=str(row['price_per_unit']))
        pdf.cell(w=40, h=10, border=1, ln=1, txt=str(row['total_price']))
    total_sum = df['total_price'].sum()
    pdf.set_font("Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=10, border=1, txt="")
    pdf.cell(w=50, h=10, border=1, txt="")
    pdf.cell(w=40, h=10, border=1, txt="")
    pdf.cell(w=30, h=10, border=1, txt="")
    pdf.cell(w=40, h=10, border=1, txt=str(total_sum), ln=1)
    pdf.set_font("Times", size=16)
    pdf.cell(w=30,h=10,txt=f"Thawani Corp.")
    pdf.image("images.png", w=30, h=20, x=50)
    pdf.output(f"PDF/{filename}.pdf")