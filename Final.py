import pandas as pd
import json
import re



def lookup_columns_in_current_sheet(excel_file_path, json_file_path):
    sheet_names = pd.ExcelFile(excel_file_path).sheet_names

    with open(json_file_path) as f:
        json_data = json.load(f)
        lookup_columns = json_data["columns"]

    all_lookup_columns_df = pd.DataFrame()

    for sheet in sheet_names:
        sheet_df = pd.read_excel(excel_file_path, sheet_name=sheet)

        lookup_columns_in_current_sheet = [column for column in lookup_columns if column in sheet_df.columns]

        for column in lookup_columns_in_current_sheet:
            all_lookup_columns_df[column] = sheet_df[column]

    missing_lookup_columns = list(set(lookup_columns) - set(all_lookup_columns_df.columns))

    if not missing_lookup_columns:
        print("The Excel File contains all lookup columns.")
    else:
        print("The DataFrame does not contain the following lookup columns:")
        print(missing_lookup_columns)
        excel_df=validate_excel_file(excel_file_path, json_file_path)
        return excel_df

    return all_lookup_columns_df 

def compare_excel_with_json(excel_df, json_file_path):
    with open(json_file_path, 'r') as json_file:
        json_data = json.load(json_file)

 

    json_columns = json_data['columns']
    json_datatypes = json_data['datatypes']

 

    mismatched_columns = []
    
   
    excel_df = excel_df.convert_dtypes(convert_string='infer')
     
    

 

    for col in excel_df.columns:
        if col in json_columns:
            excel_data_type = str(excel_df[col].dtype)
            json_data_type = json_datatypes[col]

 

            if excel_data_type != json_data_type:
                mismatched_columns.append((col, excel_data_type, json_data_type))

 

    return mismatched_columns

 

def replace_special_characters(column_name):
    # Define a regular expression pattern to match special characters
    pattern = r'[^\w\d]+'
    return re.sub(pattern, '', column_name)


 

 

def validate_excel_file(excel_file_path, json_file_path):
    def check_transposed(df, lookup_df):
        for row in range(len(df)):
            if set(df.iloc[row]) == set(lookup_df.columns):
                return True
        return False
    try:
        # Step 1: Read multiple sheets from the Excel file
        xls = pd.ExcelFile(excel_file_path)
        sheets = xls.sheet_names


 

    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")
    first_sheet_columns = None
    lookup_columns = set()
    for sheet_name in sheets:
        try:
            # Read the current sheet and lookup file
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            df_header_none=pd.read_excel(excel_file_path, sheet_name=sheet_name,header=None)
            df.dropna(how='all', inplace=True)
            df_header_none.dropna(how='all', inplace=True)
            with open(json_file_path, 'r') as json_file:
                lookup_data = json.load(json_file)

            # Handle nested dictionaries if necessary
            if isinstance(lookup_data, dict):
                lookup_df = pd.DataFrame.from_dict(lookup_data, orient='index').T
            else:
                lookup_df = pd.DataFrame(lookup_data)

            lookup_df = lookup_data['columns']

            print("Lookup Columns")
            print(lookup_df)
            # Step 2: Checked for lookup file column names in the current sheet
            missing_columns = [col for col in lookup_df if col not in df.columns]
            additional_columns = [col for col in df.columns if col not in lookup_df]
            lookup_columns.update(lookup_df)
            if  missing_columns and  df.shape[0] < df.shape[1]:
                df_header_none.dropna(how='all', inplace=True)

                transposed_df = df_header_none.transpose()
                transposed_df.columns = transposed_df.iloc[0]
                transposed_df.drop(0, inplace=True)
                transposed_columns = set(transposed_df.columns)
              


 

                # for checking special characters
                new_column_names = []
                for col in transposed_columns:
                    new_column_name = replace_special_characters(col)
                    new_column_names.append(new_column_name)
                    
                if len(new_column_names) != len(transposed_columns):
                    raise ValueError("Number of new column names doesn't match the number of existing columns.")
                transposed_columns=new_column_names

                if transposed_columns==lookup_columns:
                    print("this")
                    print(transposed_df)
                    print("Both Lookup and transposed tables are matched")
                    return(transposed_columns)
                elif len(transposed_columns)>=len(lookup_columns) or len(transposed_columns)<=len(lookup_columns):
                   
                    missing_columns = [col for col in lookup_df if col not in transposed_columns]
                    additional_columns=[col for col in transposed_columns if col not in lookup_df]

                    if additional_columns:
                        print("Additional columns present in Excel File is ",additional_columns)   

 

                    if missing_columns:
                      print("Missing columns in Excel file is")
                      print(missing_columns)

                    return(transposed_df)
                    return(lookup_df)
                    print("transposed",transposed_df)
            else:
                     # DOne this because while checking for missing columns in excel file 
                     df = pd.read_excel(excel_file_path, sheet_name=sheet_name,header=None,index_col=None)
                     df.dropna(how="all",inplace=True)
                    
                     #because the first row will be eliminate
                     df.columns=df.iloc[0]
                     
                     #Checking for Special Characters
                     new_column_names = []
                     for col in df.columns:
                        new_column_name = replace_special_characters(col)
                        new_column_names.append(new_column_name)
                     if len(new_column_names) != len(df.columns):
                        raise ValueError("Number of new column names doesn't match the number of existing columns.")
                   
                     missing_columns1 = [col for col in lookup_df if col not in df.columns]
                     additional_columns1 = [col for col in df.columns if col not in lookup_df]
                     if missing_columns1:
                         print("Missing columns")
                         print(missing_columns1)
                     if additional_columns1 : 
                         print("additional_columns")
                         print(additional_columns1)
                         print(df)
                     return df


            if first_sheet_columns is None:
                first_sheet_columns = set(df.columns)

            if set(df.columns) != first_sheet_columns:
                missing_cols_in_subsequent = first_sheet_columns - set(df.columns)
                print(f"Columns missing in sheet '{sheet_name}': {missing_cols_in_subsequent}")

 

        except Exception as e:
            raise Exception(f"Error in sheet '{sheet_name}': {e}")
        return sheet_name

    print("all are validated")  
    print("All sheets have been validated and values printed successfully.")


 

# Usage example
excel_file_path = r'C:\Users\preethi.s\Downloads\tests.xlsx'
json_file_path = r'C:\Users\preethi.s\Documents\RDBMS\schema.json'

 
all_lookup_columns_df = lookup_columns_in_current_sheet(excel_file_path, json_file_path)

print(all_lookup_columns_df)

xls = pd.ExcelFile(excel_file_path)
sheets = xls.sheet_names
mismatched_columns = compare_excel_with_json(all_lookup_columns_df, json_file_path)
if not mismatched_columns:
    print("All column data types match the JSON object.")
else:
    print("*****Mismatched column data types*****:")
    for col, excel_data_type, json_data_type in mismatched_columns:
        print(f"Column '{col}' has data type '{excel_data_type}', expected '{json_data_type}'.")

 


print("All Sheets are validated Succesfully")
print("Converting all values to JSON format")
json_data = all_lookup_columns_df.to_json(orient='records')
print(json_data)
