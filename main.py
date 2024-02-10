from pathlib import Path
import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Getting the list of files
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1, border=0, align="L")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Date {date}", ln=2, border=0, align="L")

    # Reading the excel
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    # Add the table header
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], ln=0, border=1, align="L")
    pdf.cell(w=65, h=8, txt=columns[1], ln=0, border=1, align="L")
    pdf.cell(w=35, h=8, txt=columns[2], ln=0, border=1, align="L")
    pdf.cell(w=30, h=8, txt=columns[3], ln=0, border=1, align="L")
    pdf.cell(w=30, h=8, txt=columns[4], ln=1, border=1, align="L")

    # Add the table content
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), ln=0, border=1, align="L")
        pdf.cell(w=65, h=8, txt=row["product_name"], ln=0, border=1, align="L")
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), ln=0, border=1, align="L")
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), ln=0, border=1, align="L")
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=1, border=1, align="L")

    # Add the total row to table
    invoice_total = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt="", ln=0, border=1, align="L")
    pdf.cell(w=65, h=8, txt="", ln=0, border=1, align="L")
    pdf.cell(w=35, h=8, txt="", ln=0, border=1, align="L")
    pdf.cell(w=30, h=8, txt="", ln=0, border=1, align="L")
    pdf.cell(w=30, h=8, txt=str(invoice_total), ln=1, border=1, align="L")

    pdf.cell(w=0, h=8, txt="", ln=1, border=0, align="L")

    # Adding the final lines
    pdf.set_font(family="Times", style='B', size=16)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f"The total due amount is {invoice_total} Euros", ln=1, border=0, align="L")

    pdf.cell(w=30, h=8, txt="PythonHow", ln=0, border=0, align="L")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")