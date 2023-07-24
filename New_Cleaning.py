import pandas as pd
from openpyxl import load_workbook

def validate_excel_file(filename, lookup_file):
    def check_transposed(df, lookup_df):
        for row in range(len(df)):
            if set(df.iloc[row]) == set(lookup_df.columns):
                return True
        return False

    try:
        # Step 1: Read multiple sheets from the Excel file
        xls = pd.ExcelFile(filename)
        sheets = xls.sheet_names
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")

    first_sheet_columns = None
    lookup_columns = set()

    for sheet_name in sheets:
        try:
            # Read the current sheet and lookup file
            df = pd.read_excel(filename, sheet_name=sheet_name)
            lookup_df = pd.read_excel(lookup_file)

            # Step 2: Checked for lookup file column names in the current sheet
            missing_columns = [col for col in lookup_df.columns if col not in df.columns]
            additional_columns = [col for col in df.columns if col not in lookup_df.columns]
            lookup_columns.update(lookup_df.columns)

            if missing_columns:
                # Step 3: Checked whether column names are present in subsequent rows or transposed
                if check_transposed(df, lookup_df):
                    raise Exception(f"Columns are transposed in sheet '{sheet_name}'.")
                else:
                    missing_columns_str = ", ".join(missing_columns)
                    raise Exception(f"Columns {missing_columns_str} are not found in sheet '{sheet_name}'.")

            # Step 4: Validate data for missing input columns
            missing_input_columns = [col for col in lookup_df.columns if col not in df.columns]
            if missing_input_columns:
                # Notified for validation failures
                missing_input_columns_str = ", ".join(missing_input_columns)
                raise Exception(f"Missing input columns in sheet '{sheet_name}': {missing_input_columns_str}")

            # Print excel sheet values
            print(f"Values in sheet '{sheet_name}':")
            print(df)
            #saved the first comparison
            if first_sheet_columns is None:
                first_sheet_columns = set(df.columns)

            # Compare the columns of the current sheet with the first sheet
            if set(df.columns) != first_sheet_columns:
                missing_cols_in_subsequent = first_sheet_columns - set(df.columns)
                print(f"Columns missing in sheet '{sheet_name}': {missing_cols_in_subsequent}")

        except Exception as e:
            raise Exception(f"Error in sheet '{sheet_name}': {e}")

    # Notify users of additional columns in data not present in lookup
    additional_columns_in_data = [col for col in first_sheet_columns if col not in lookup_columns]
    if additional_columns_in_data:
        print(f"Additional columns in data not present in lookup: {additional_columns_in_data}")

    print("All sheets have been validated and values printed successfully.")
validate_excel_file(r'C:\Users\preethi.s\Downloads\tests.xlsx', r'C:\Users\preethi.s\Downloads\test_lookup.xlsx')
