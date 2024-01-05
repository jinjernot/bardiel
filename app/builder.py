import pandas as pd
from app.word import *

def create_sheet(xlsx_file):
    """Builds a sheet"""

    # Read Excel
    df = pd.read_excel(xlsx_file, sheet_name='Sheet1', skiprows=3)

    # Remove rows
    df = df.drop(index=range(0, 30))

    # Remove rows where "Container Group" contains "Messaging"
    df = df[~df["Container Group"].str.contains("Messaging", na=False)]

    # Remove rows where "Container Group" contains "Messaging"
    df = df[~df["Container Group"].str.contains("Facets", na=False)]

    # Remove rows where "Container Group" contains "Messaging"
    df = df[~df["Container Group"].str.contains("Core Information", na=False)]

        # Remove rows where "Container Group" contains "Messaging"
    df = df[~df["Container Group"].str.contains("Metadata", na=False)]

    # Iterate through each row and update "Series Value" with unique values
    for index, row in df.iterrows():
        unique_values_set = set(row[7:].dropna())  # Remove NaN if present
        unique_values_str = ', '.join(map(str, unique_values_set))

        # Remove duplicates within the row in "Series Value" column
        series_values_list = list(set(str(row["Series Value"]).split(', ')))
        series_values_list.extend(unique_values_set)
        df.at[index, "Series Value"] = ', '.join(map(str, set(series_values_list)))

    # Replace "#Intentionally Left Blank#" with NaN in "Series Value" column
    df["Series Value"] = df["Series Value"].replace("#Intentionally Left Blank#", float('nan'), regex=True)

    # Remove rows where "Series Value" is NaN or "#Intentionally Left Blank#"
    df = df.dropna(subset=["Series Value"])
    df = df[~((df["Series Value"] == "#Intentionally Left Blank#") & df.iloc[:, 7:].isna().all(axis=1))]

    # Remove rows where "Series Value" is the string "nan"
    df = df[df["Series Value"] != "nan"]


   # Keep only "Container Name" and "Series Value" columns
    df = df[["Container Name", "Series Value"]]

    # Save Excel file
    excel_file = 'data.xlsx'
    df.to_excel(excel_file, index=False)

    # Convert Excel to Word
    word_output_file = 'data.docx'
    excel_to_word(excel_file, word_output_file)
