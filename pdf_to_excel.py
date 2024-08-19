import pdfplumber
import pandas as pd
import click

def extract_tables_from_pdf(pdf_path):
    # Initialize a list to hold the tables
    all_tables = []

    # Open the PDF file
    with pdfplumber.open(pdf_path) as pdf:
        # Iterate through all the pages
        for page in pdf.pages:
            # Extract tables from the page
            tables = page.extract_tables()
            all_tables.extend(tables)

    return all_tables

def tables_to_excel(tables, excel_path):
    # Initialize an Excel writer
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            # Convert table to DataFrame
            df = pd.DataFrame(table[1:], columns=table[0])
            # Write DataFrame to Excel
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

    print(f"PDF content has been successfully written to {excel_path}")

@click.command()
@click.argument('pdf_path')
@click.argument('excel_path')
def pdf_to_excel(pdf_path, excel_path):
    # Extract tables from the PDF
    tables = extract_tables_from_pdf(pdf_path)
    # Save tables to an Excel file
    tables_to_excel(tables, excel_path)

if __name__ == '__main__':
    pdf_to_excel()
    