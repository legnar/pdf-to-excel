import tabula
import pandas as pd
import os


def pdf_to_excel(pdf_file_path, excel_file_path):
    # Read PDF file
    tables = tabula.read_pdf(pdf_file_path, pages='all')

    # Write each table to a separate sheet in the Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Sheet{i+1}')


# pdf_to_excel('pdf.pdf', 'path_to_excel_file.xlsx')
PDF_FILES_DIR = 'pdf_files'
EXCEL_FILES_DIR = 'excel_files'

for filename in os.listdir(PDF_FILES_DIR):
    pdfpath = os.path.join(PDF_FILES_DIR, filename)
    filename = filename.split('.')[-2] + '.xlsx'
    excelpath = os.path.join(EXCEL_FILES_DIR, filename)
    pdf_to_excel(pdfpath, excelpath)
    print(f'Converted {filename}.')