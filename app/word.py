from docx import Document
import pandas as pd
from app.table import table_column_widths
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def excel_to_word(df, word_file):
    # Create a Word document
    doc = Document()

    # Add the DataFrame as a table to the Word document
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])

    table_column_widths(table, (Inches(2), Inches(5.5),))

    # Iterate through DataFrame columns and values
    for i, column in enumerate(df.columns):
        table.cell(0, i).text = column  # Set column headers
        for j in range(df.shape[0]):
            value = df.iloc[j, i]
            cell = table.cell(j + 1, i)
            cell.text = str(value)
            # Set font size for each cell in the table
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(6)  # Set font size to 6 points
            # Reduce spacing within cell
            cell.paragraphs[0].paragraph_format.space_before = Pt(0)
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)


    # Set font size and spacing for table header
    for cell in table.rows[0].cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(6)  # Set font size to 6 points
        # Add spacing to table header
        cell.paragraphs[0].paragraph_format.space_before = Pt(6)
        cell.paragraphs[0].paragraph_format.space_after = Pt(6)

    # Set table width to match entire page width
    table.autofit = False
    section = doc.sections[-1]
    table_width = section.page_width - section.left_margin - section.right_margin
    table.width = table_width

    # Set margins
    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.5)  # Set left margin to 0.5 inches
        section.right_margin = Inches(0.5)  # Set right margin to 0.5 inches
        section.top_margin = Inches(0.5)  # Set top margin to 0.5 inches
        section.bottom_margin = Inches(0.5)  # Set bottom margin to 0.5 inches

    # Save the Word document
    doc.save(word_file)