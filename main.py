import pandas as pd
"""
glob function is used to search for files that match a specific file pattern or name 
"""
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    # stem extracts the main filename without extension
    filename = Path(filepath).stem
    # Extracts only the first item of list
    invoice_nmbr = filename.split("-")[0]
    """
    We can get invoice_number and file name by using slice method
    """
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=0, h=12, txt=f"Invoice nr.{invoice_nmbr}", align="L", ln=1)
    pdf.output(f"PDFs/{filename}.pdf")
