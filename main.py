# installed openpyxl to open Excel files
# installed Fpdf to create PDF  files
# to write txt in PDF we use cell() method
# import Path to get file location(path) of the xlsx file
# with .stem we extract only the string
# glob module-> global, is a function that's used to search for files that match a specific file pattern or name.
# pandas to read Excel file
# filepath creates a list of filepaths
# ln indicates a break line should be created

import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="p", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    pdf.output(f"PDFs/{filename}.pdf")
