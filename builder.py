import pandas as pd

def createSheet(xlsx_file):
    # Read Excel file into a DataFrame
    df = pd.read_excel(xlsx_file, sheet_name='Sheet1', skiprows=3)

    # Remove rows 5 to 31
    df = df.drop(index=range(0, 30))
    
    # Append values from column 7 to the end
    selected_columns = pd.concat([df[["Container Name", "Container Group"]], df.iloc[:, 7:]], axis=1)

    # Remove rows with NaN values in selected columns
    selected_columns = selected_columns.dropna()

    # Add a filter to select only rows where "Container Group" contains "Technical Specifications"
    selected_columns = selected_columns[selected_columns["Container Group"].str.contains("Technical Specifications", na=False)]

    # Save the selected columns data to a new Excel file
    output_file = 'output_selected_columns.xlsx'
    selected_columns.to_excel(output_file, index=False)