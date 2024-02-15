from app.builder import create_sheet

import glob

def main():
    """Load the xlsx files and create sheets"""

    folder_path = "./docs/" 
    xlsx_files = glob.glob(folder_path + "*.xlsx")
    
    for xlsx_file in xlsx_files:
        create_sheet(xlsx_file)

if __name__ == "__main__":
    main() 