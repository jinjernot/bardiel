from openpyxl import load_workbook
from docx import Document
import pandas as pd

def excel_to_word(excel_file, word_file):
    # Load Excel workbook
    book = load_workbook(excel_file)

    # Create a Word document
    doc = Document()

    # Iterate through each sheet in the Excel workbook
    for sheet_name in book.sheetnames:
        # Add a heading with the sheet name
        doc.add_heading(sheet_name, level=1)

        # Get the data from the Excel sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # Add the DataFrame as a table to the Word document
        doc.add_table(df.shape[0] + 1, df.shape[1]).style = 'Table Grid'
        for i, column in enumerate(df.columns):
            doc.tables[-1].cell(0, i).text = column
            for j, value in enumerate(df[column]):
                doc.tables[-1].cell(j + 1, i).text = str(value)

        # Add a page break between sheets
        doc.add_page_break()

    # Save the Word document
    doc.save(word_file)