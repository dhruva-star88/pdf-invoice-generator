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
    # create pdf document for each fila-path
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    # Adds page for each pdf
    pdf.add_page()
    # stem extracts the main filename without extension
    filename = Path(filepath).stem
    # Extracts only the first item of list
    invoice_nmbr = filename.split("-")[0]
    date = filename.split("-")[1]
    """
    We can get invoice_number and file name by using slice method
    """
    # Invoice number
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=0, h=12, txt=f"Invoice nr.{invoice_nmbr}", align="L", ln=1)
    # Date
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=0, h=12, txt=f"Date: {date}", align="L", ln=1)
    pdf.ln(10)

    # Read dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)
    columns = [column.replace("_", " ").title() for column in columns]
    print(columns)

    # Add Headers
    pdf.set_font(family='Times', size=12, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=columns[0], border=1)
    pdf.cell(w=55, h=10, txt=columns[1], border=1)
    pdf.cell(w=40, h=10, txt=columns[2], border=1)
    pdf.cell(w=30, h=10, txt=columns[3], border=1)
    pdf.cell(w=30, h=10, txt=columns[4], border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=str(row["product_id"]), border=1)
        pdf.cell(w=55, h=10, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), border=1, ln=1)

    # Add Total Price
    """
    price_list = []
    for index, row in df.iterrows():
        price = row["total_price"]
        price_list.append(price)
        total_price = sum(price_list)
     """
    total_price = df["total_price"].sum()

    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=55, h=10, txt="", border=1)
    pdf.cell(w=40, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=str(total_price), border=1, ln=1)

    # Add total price sentence
    pdf.set_font(family='Times', size=14, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=10, txt=f"The Total Price is: {total_price}", ln=1)
    # Add Company name and logo
    pdf.set_font(family='Times', size=14, style="B")
    pdf.cell(w=27, h=10, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
