import pandas as pd
from app.word import *

import pandas as pd

def create_sheet(xlsx_file):
    """Builds a sheet"""

    # Read Excel
    df = pd.read_excel(xlsx_file, sheet_name='Sheet1', skiprows=3)

    # Remove formatting from text (e.g., italicized text)
    df = df.applymap(lambda x: x if not hasattr(x, 'font') else x.value)

    # Convert all data to string format
    df = df.astype(str)

    # Remove rows
    df = df.drop(index=range(0, 30))

    # Remove rows where "Container Group" contains specific strings
    unwanted_strings = ["Messaging", "Facets", "Core Information", "Metadata"]
    df = df[~df["Container Group"].str.contains('|'.join(unwanted_strings), na=False)]


    # Keep only "Container Name" and "Series Value" columns
    df = df[["Container Name", "Series Value"]]

    # Remove rows where "Series Value" is "nan|#Intentionally Left Blank#"
    df = df[~df["Series Value"].isin(["#Intentionally Left Blank#"])]
    df = df[~df["Series Value"].isin(["nan"])]


    
    # Save Excel file
    excel_file = 'data.xlsx'
    df.to_excel(excel_file, index=False)

    # Convert Excel to Word
    word_file = 'data.docx'
    excel_to_word(df, word_file)