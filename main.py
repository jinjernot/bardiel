from builder import createSheet

import glob

def loadxlsx():
    """Load the xlsx file"""
    
    folder_path = "./xlsx/" 
    xlsx_files = glob.glob(folder_path + "*.xlsx")
    
    for xlsx_file in xlsx_files:
        createSheet(xlsx_file)
        
def main():
    loadxlsx()

if __name__ == "__main__":
        main()  