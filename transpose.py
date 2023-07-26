import pandas as pd
from openpyxl import load_workbook

 

 

def validate_excel_file(filename, lookup_file):
        df =  pd.read_excel(filename)
        lookup_df =  pd.read_excel(lookup_file)
        lookup_columns=lookup_df.columns
        df_transposed = df.T
        new_headers = df_transposed.iloc[0]
        df_transposed.columns = new_headers
        df_transposed = df_transposed[1:]
        print(df_transposed.columns)

        if len(df.columns)>=len(lookup_df.columns) and  df_transposed.columns.equals(lookup_df.columns):
            print("first one runs")
        elif df_transposed.columns.equals(lookup_df.columns):
           print("transpose")
        elif df.columns.equals(lookup_df.columns) and  df_transposed.columns.equals(lookup_df.columns):

 

               print("loop runs")
               for col1 in df.columns:
                 for col2 in lookup_df.columns:
                   if col1 == col2:
                     print(f"Column '{col1}' exists in both files.")
                     break
                 else:
                   print(f"Additional column : '{col1}' does not exist in the lookup  file.")
        elif not df.columns.equals(lookup_df.columns) and not df_transposed.columns.equals(lookup_df.columns):
         print(df_transposed)

        
         for col1 in df_transposed.columns:
             for col2 in lookup_df.columns:
                 if col1==col2:
                     print(f"Column '{col1}' exists in both files.")
                     break
             else:
                 additional_columns = [col for col in df_transposed.columns if col not in lookup_df.columns]
                 if additional_columns:
                     print("Additional columns found:")
                     print(additional_columns)
                     print(df)

 

        else:
            additional_columns = [col for col in df.columns if col not in lookup_df.columns]
            if additional_columns:
                     print("Additional columns found:")
                     print(additional_columns)
                     print(df)
            print(df)


validate_excel_file(r'C:\Users\preethi.s\Downloads\tests.xlsx', r'C:\Users\preethi.s\Downloads\test_lookup.xlsx')
