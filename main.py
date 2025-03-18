import pandas as pd
from functions import *

def main():
    # Set option to display all rows
    pd.set_option('display.max_rows', None)
    
    # The filename in the same directory
    excel_file = 'excel_file.xlsx'
    
    # Call removeAddress and store the returned DataFrame
    modified_df = removeAddress(excel_file)
    
    # Print the entire DataFrame as a string
    print(modified_df.to_string())
    print("Total rows:", modified_df.shape[0])

if __name__ == '__main__':
    main()
