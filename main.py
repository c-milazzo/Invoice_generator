import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:

    # create PDF page
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # get filename/date
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    # input invoice info
    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=10, h=4, txt=" ", ln=1)

    # input date info
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)
    pdf.cell(w=10, h=4, txt=" ", ln=1)

    # Read data
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = [item.replace("_", " ").title() for item in df.columns]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(10, 10, 10)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add the sum for total price
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.cell(w=10, h=4, txt=" ", ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=f"The total amount due is ${total_sum}", ln=1)

    # Add company name, logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=26, h=8, txt="PythonHow ")
    pdf.image("pythonhow.png", w=8)

    pdf.output(f"PDFs/{filename}.pdf")
