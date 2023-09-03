import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    #read data
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #create Pdf page
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    #get filename/date
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    #create cell
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}")




    pdf.output(f"PDFs/{filename}.pdf")
