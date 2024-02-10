from pathlib import Path
import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1, border=0, align="L")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Date {date}", ln=1, border=0, align="L")

    pdf.output(f"PDFs/{filename}.pdf")