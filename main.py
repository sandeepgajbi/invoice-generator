import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


def generate_pdf(filepath):
    try:
        # Read Excel file
        df = pd.read_excel(filepath, sheet_name="Sheet 1")

        # Extract invoice number and date from file name
        invoice_number, invoice_date = Path(filepath).stem.split("-")

        # Create PDF object
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.add_page()

        # Set font for title
        pdf.set_font(family="Times", size=16, style="B")

        # Add invoice number and date to PDF
        pdf.cell(w=50, h=8, txt=f"Invoice number : {invoice_number}", ln=1)
        pdf.cell(w=50, h=8, txt=f"Date : {invoice_date}", ln=1)

        # Set font for table headers and data
        pdf.set_font(family="Times", size=10, style="B")

        # Add table headers to PDF
        column = df.columns
        columns = [item.replace("_", " ").title() for item in column]
        col_widths = [pdf.get_string_width(col) + 18 for col in columns]
        row_height = 12
        for col, col_width in zip(columns, col_widths):
            pdf.cell(w=col_width, h=row_height, txt=col, border=1)
        pdf.ln()

        # Add data rows to PDF
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        for index, row in df.iterrows():
            for col, col_width in zip(column, col_widths):
                pdf.cell(w=col_width, h=row_height, txt=str(row[col]), border=1)
            pdf.ln()

        # data row to PDF with Total price
        total_sum = df["total_price"].sum()
        for col, col_width in zip(columns, col_widths):
            if col == "Total Price":
                pdf.cell(w=col_width, h=row_height, txt=str(total_sum), border=1)
            else:
                pdf.cell(w=col_width, h=row_height, txt="", border=1)
        pdf.ln()

        # Add total price to the next line
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=50, h=row_height, txt=f"The total price is : {total_sum}")

        # Output PDF file
        pdf.output(f"PDFs/{invoice_number}.pdf")

    except Exception as e:
        print(f"Error generating PDF for '{filepath}': {e}")


def main():
    # Get list of Excel files in "invoices" directory
    filepaths = glob.glob("invoices/*.xlsx")

    # Generate PDF for each Excel file
    for filepath in filepaths:
        generate_pdf(filepath)


if __name__ == "__main__":
    main()