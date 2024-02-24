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

    df_skus = df.copy()
    

    # Extracting values from "Container Name" column
    container_name_values = df_skus["Container Name"].values

    # Extracting values starting from column 7
    values_from_column_7 = df_skus.iloc[:, 7:].values

    # Create a new DataFrame with the extracted values
    new_df = pd.DataFrame({'Container Name': container_name_values})

    # Get the column names from the original DataFrame starting from column 7
    original_column_names = df_skus.columns[7:]

    # Iterate over the remaining columns and add them to the new DataFrame
    for col_name, col_values in zip(original_column_names, values_from_column_7.T):
        new_df[col_name] = col_values

    # Remove rows with "nan" values
    new_df = new_df[~new_df.isin(['nan']).any(axis=1)]

    # Printing the new DataFrame


    print(new_df)

    # Keep only "Container Name" and "Series Value" columns
    df = df[["Container Name", "Series Value"]]

    # Remove rows where "Series Value" is "nan|#Intentionally Left Blank#"
    df = df[~df["Series Value"].isin(["#Intentionally Left Blank#"])]
    df = df[~df["Series Value"].isin(["nan"])]
    
    # Save Excel file
    excel_file = 'data.xlsx'
    df.to_excel(excel_file, index=False)
    new_df = new_df.to_excel("skus.xlsx", index=False)

    # Convert Excel to Word
    word_file = 'data.docx'
    excel_to_word(df, word_file)
