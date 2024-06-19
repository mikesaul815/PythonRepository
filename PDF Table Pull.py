import pdfplumber
from openpyxl import Workbook

def extract_table_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        # Extract text from all pages
        text = ''
        for page in pdf.pages:
            text += page.extract_text()

        # Extract tables from all pages
        tables = []
        for page in pdf.pages:
            tables.extend(page.extract_tables())

    return text, tables

def write_to_excel(text, tables, excel_path):
    wb = Workbook()
    ws = wb.active

    # Write extracted text to the first cell
    ws['A1'] = text

    # Write each table to Excel
    row_offset = 3  # Offset to leave space for the extracted text
    for idx, table in enumerate(tables):
        for row in table:
            ws.append(row)
        ws.append([])  # Add an empty row between tables
        row_offset += len(table) + 1

    wb.save(excel_path)

if __name__ == "__main__":
    pdf_path = "filepath/file.pdf"  # Replace with your PDF file path
    excel_path = "filepath/file.xlsx"  # Replace with desired Excel file path

    text, tables = extract_table_from_pdf(pdf_path)
    write_to_excel(text, tables, excel_path)
    print("Table information extracted from PDF and written to Excel successfully.")
