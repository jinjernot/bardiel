import pandas as pd

def createSheet(xlsx_file):
    """Builds a sheet"""

    # Read Excel
    df = pd.read_excel(xlsx_file, sheet_name='Sheet1', skiprows=3)

    # Remove rows
    df = df.drop(index=range(0, 30))
    
    # Append values from column 7 to the end
    selected_columns = pd.concat([df[["Container Name", "Container Group", "Series Value"]], df.iloc[:, 7:]], axis=1)

    # Remove rows with NaN
    selected_columns = selected_columns.dropna()

    # Select rows where "Container Group" contains "Technical Specifications"
    #selected_columns = selected_columns[selected_columns["Container Group"].str.contains("Technical Specifications", na=False)]

    # Remove "Container Group" column
    #selected_columns = selected_columns.drop("Container Group", axis=1)

    # Save Excel file
    output_file = 'data.xlsx'
    selected_columns.to_excel(output_file, index=False)