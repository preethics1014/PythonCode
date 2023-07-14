import pandas as pd
import sqlite3

def clean_column_names(df):
    df.columns = df.columns.str.replace(' ', '_')
    df.columns = df.columns.str.replace('[!@#$%^&*()_]', '')
    return df
def read_excel_file(filepath):
    xls = pd.ExcelFile(filepath)
    print(xls.sheet_names)

    for sheet_name in xls.sheet_names:
        print(sheet_name)
        sql_table = 'tbl_' + sheet_name
        print('Table -', sql_table)
       
  
        # Read the sheet into a DataFrame
        df = pd.read_excel(filepath, sheet_name=sheet_name, index_col=None, header=None)
                # Check if the row is transposed
        if df.empty:
            print(f"Sheet '{sheet_name}' is empty. Skipping...")
            continue
         # Ignore the first row if it is not present
        if pd.isnull(df.iloc[0]).all():
            df = df[1:]

      

        # Ignore the row if it is transposed
        if df.shape[0] < df.shape[1]:
            df = df.transpose()
       
        print(df.head(10))
        


filepath = r'C:\Users\preethi.s\Downloads\productcopy.xlsx'
print(filepath)
read_excel_file(filepath)
