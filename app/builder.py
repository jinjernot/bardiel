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

    # Iterate through each row and update "Series Value" with unique values
    for index, row in df.iterrows():
        unique_values_set = set(row[7:].dropna())  # Remove NaN if present
        unique_values_str = '|'.join(map(str, unique_values_set))  # Changed comma to pipe character

        # Remove duplicates within the row in "Series Value" column
        series_values_list = list(set(str(row["Series Value"]).split('|')))  # Changed comma to pipe character
        series_values_list.extend(unique_values_set)
        df.at[index, "Series Value"] = '|'.join(map(str, set(series_values_list)))  # Changed comma to pipe character

    # Keep only "Container Name" and "Series Value" columns
    df = df[["Container Name", "Series Value"]]

    # Remove rows where "Series Value" is "nan|#Intentionally Left Blank#"
    df = df[~df["Series Value"].isin(["nan|#Intentionally Left Blank#"])]
    df = df[~df["Series Value"].isin(["#Intentionally Left Blank#|nan"])]
    df = df[~df["Series Value"].isin(["nan"])]

    # Clean up "Series Value" column
    def clean_series_name(series_name):
        series_name = series_name.replace("nan|", "").replace("|nan", "").replace("|#Intentionally Left Blank#", "").replace("#Intentionally Left Blank#|", "").replace("nan", "")
        return series_name

    df["Series Value"] = df["Series Value"].apply(clean_series_name)

    # Remove rows where "Series Value" is empty
    df = df[df["Series Value"].notna()]
    df["Series Value"] = df["Series Value"].str.replace('|', ' | ')
    
    # Save Excel file
    excel_file = 'data.xlsx'
    df.to_excel(excel_file, index=False)

    # Convert Excel to Word
    word_file = 'data.docx'
    excel_to_word(df, word_file)