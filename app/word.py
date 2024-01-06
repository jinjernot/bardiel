from openpyxl import load_workbook
from docx.shared import Inches
from docx import Document
import pandas as pd

def excel_to_word(excel_file, word_file):
    # Load Excel workbook
    book = load_workbook(excel_file)

    # Create a Word document
    doc = Document()

    # Get the active sheet from the Excel workbook
    sheet = book.active

    # Get the data from the Excel sheet
    data = sheet.values
    columns = next(data)
    df = pd.DataFrame(data, columns=columns)

    # Add the DataFrame as a table to the Word document
    table = doc.add_table(rows=df.shape[0], cols=df.shape[1])

                
    # Iterate through DataFrame columns and values
    for i, column in enumerate(df.columns):
        for j in range(df.shape[0]):
            value = df.iloc[j, i]
            doc.tables[-1].cell(j, i).text = str(value)

    # Set margins
    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.5)  # Set left margin to 0.5 inches
        section.right_margin = Inches(0.5)  # Set right margin to 0.5 inches
        section.top_margin = Inches(0.5)  # Set top margin to 0.5 inches
        section.bottom_margin = Inches(0.5)  # Set bottom margin to 0.5 inches

    # Add a page break between sheets
    doc.add_page_break()

    # Save the Word document
    doc.save(word_file)
