import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

# Helper function to check if a file exists
def file_exists(filename):
    return os.path.exists(filename)

# Helper function to get valid input for min/max range
def get_float_input(prompt, default=0.0):
    try:
        value = input(prompt)
        if value == "":
            return default
        return float(value)
    except ValueError:
        print("Invalid input. Please input a valid number.")
        return get_float_input(prompt, default)

# Helper function to get valid percentage input (1-100)
def get_percentage_input(prompt):
    try:
        value = int(input(prompt))
        if 1 <= value <= 100:
            return value
        else:
            print("Please input an integer in the range of 1 to 100.")
            return get_percentage_input(prompt)
    except ValueError:
        print("Please input an integer in the range of 1 to 100.")
        return get_percentage_input(prompt)

# Helper function to get rounding choice
def get_rounding_choice():
    try:
        choice = int(input(
            "Please select how you want to round the value:\n"
            " 1. Down to the nearest integer\n"
            " 2. Up to the nearest integer\n"
            " 3. To 1 decimal point\n"
            " 4. To 2 decimal points:\n "
            " Selection(1-4): "
        ))
        if choice in [1, 2, 3, 4]:
            return choice
        else:
            print("Please select a valid rounding option (1-4).")
            return get_rounding_choice()
    except ValueError:
        print("Please select a valid rounding option (1-4).")
        return get_rounding_choice()

# Function to round a number based on user choice
def round_value(value, rounding_choice):
    if rounding_choice == 1:
        return int(value)
    elif rounding_choice == 2:
        return int(value) + (1 if value % 1 != 0 else 0)
    elif rounding_choice == 3:
        return round(value, 1)
    elif rounding_choice == 4:
        return round(value, 2)

# Function to modify values based on the case
def modify_values(df, min_val, max_val, case):
    # Check if 'Αξία' exists in the DataFrame
    if 'Αξία' not in df.columns:
        print("Error: The 'Αξία' column was not found in the file. Please make sure the file has the correct format.")
        return df
    
    # Exclude the last row from the selection (assuming the last row contains sums)
    product_df = df.iloc[:-1]  # All rows except the last one (which has the sums)

    # Select rows within the min/max range for column 'Αξία'
    selected_rows = product_df[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val)]
    print(f"Selected {len(selected_rows)} rows in the range Min: {min_val}, Max: {max_val}")

    if case == 1:  # Reduce quantity of product (Apply same logic to both columns F and H)
        percentage = get_percentage_input("Please input the percentage to reduce the quantity of the products: ")
        rounding_choice = get_rounding_choice()

        # Reduce 'Ποσ.1' (column F) and 'Ποσ.2' (column H) by the given percentage and round accordingly
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Ποσ.1'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Ποσ.1'
        ].apply(lambda x: round_value(x * (1 - percentage / 100), rounding_choice)).astype(float)

        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Ποσ.2'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Ποσ.2'
        ].apply(lambda x: round_value(x * (1 - percentage / 100), rounding_choice)).astype(float)

        # Recalculate 'Αξία' as 'Ποσ.1' * 'Τιμή κόστους' for all affected rows
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Αξία'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val)
        ].apply(lambda row: row['Ποσ.1'] * row['Τιμή κόστους'], axis=1)

    elif case == 2:  # Reduce per-unit value of product
        percentage = get_percentage_input("Please input the percentage to reduce the per-unit value of the products: ")
        rounding_choice = get_rounding_choice()
        
        # Reduce 'Τιμή κόστους' by the given percentage and round accordingly
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Τιμή κόστους'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Τιμή κόστους'
        ].apply(lambda x: round_value(x * (1 - percentage / 100), rounding_choice)).astype(float)

        # Recalculate 'Αξία' as 'Ποσ.1' * 'Τιμή κόστους'
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Αξία'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val)
        ].apply(lambda row: row['Ποσ.1'] * row['Τιμή κόστους'], axis=1)

    elif case == 3:  # Reduce total value of product (both quantity and unit price)
        percentage = get_percentage_input("Please input the percentage to reduce both the quantity and per-unit price of the products: ")
        rounding_choice = get_rounding_choice()
        
        # Reduce both 'Ποσ.1' and 'Τιμή κόστους' by the given percentage and round accordingly
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), ['Ποσ.1', 'Τιμή κόστους']] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), ['Ποσ.1', 'Τιμή κόστους']
        ].apply(lambda row: (round_value(row['Ποσ.1'] * (1 - percentage / 100), rounding_choice),
                             round_value(row['Τιμή κόστους'] * (1 - percentage / 100), rounding_choice)), axis=1).apply(pd.Series)

        # Recalculate 'Αξία' as 'Ποσ.1' * 'Τιμή κόστους'
        product_df.loc[(product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val), 'Αξία'] = product_df.loc[
            (product_df['Αξία'] >= min_val) & (product_df['Αξία'] <= max_val)
        ].apply(lambda row: row['Ποσ.1'] * row['Τιμή κόστους'], axis=1)

    # Recalculate the sums of 'Ποσ.1', 'Ποσ.2', 'Τιμή κόστους', and 'Αξία' (ignoring the last row)
    total_pos1 = product_df['Ποσ.1'].sum()
    total_pos2 = product_df['Ποσ.2'].sum()  # Summing column H as requested
    total_cost = product_df['Τιμή κόστους'].sum()
    total_value = product_df['Αξία'].sum()

    # Update the last row with the recalculated sums
    df.loc[len(df)-1, 'Ποσ.1'] = total_pos1
    df.loc[len(df)-1, 'Ποσ.2'] = total_pos2  # Writing the sum of column H
    df.loc[len(df)-1, 'Τιμή κόστους'] = total_cost
    df.loc[len(df)-1, 'Αξία'] = total_value
    
    return df


# Function to copy formatting from the original file
def copy_formatting(original_file, new_file):
    original_wb = load_workbook(original_file)
    new_wb = load_workbook(new_file)
    
    original_ws = original_wb.active
    new_ws = new_wb.active

    # Copy font and other styles from the original file to the new file
    for row in original_ws.iter_rows(min_row=1, max_row=original_ws.max_row, min_col=1, max_col=original_ws.max_column):
        for cell in row:
            new_cell = new_ws[cell.coordinate]
            new_cell.font = Font(
                name=cell.font.name, 
                size=cell.font.size, 
                bold=cell.font.bold, 
                italic=cell.font.italic, 
                underline=cell.font.underline, 
                strike=cell.font.strike
            )

    new_wb.save(new_file)

# CLI for the Excel Modifier Tool
def main():
    print("EXCEL FILE MODIFIER (V1)\n"
          "                        ")
    print("Program rules:\n"
          "                ")
    print("1. The program will always copy  the first 8 rows of a file selected and paste them in thew new file once its created\n" 
          "2. The program will always ignore the first 8 rows and start implementing the row range selection from row 9 and on\n" 
          "3. The program will recalculate the sum of columns F I and J in all iterations, ignoring the existing values of the original file's last row\n"
          "(this prevents the sum calculation from including the previous sum value)\n")
    
    while True:
        action = input("Please select your next action:\n 1. Choose a file\n 2. Exit:\n Selection(1/2): ")
        if action == '1':
            while True:
                filename = input("Please place your file in the same directory as this script and give the full name of your file (e.g., filename.xlsx): ")
                if file_exists(filename):
                    df = pd.read_excel(filename, skiprows=6)  # Ignore the first 6 rows, starting from row 7
                    
                    print(f"The file you selected is: {filename} ({os.path.getsize(filename)} bytes)")
                    next_action = input("Please select your next action:\n 1. Modify values\n 2. Exit:\n Selection(1-2): ")
                    if next_action == '1':
                        print("You will now have to specify what rows you want to modify,\nthis is done by providing a min and a max (a range of product value) from the  columns J\n")
                        min_value = get_float_input("Please input your min value (press enter for 0): ", 0.0)
                        max_value = get_float_input("Please input your max value: ")
                        print(f"Min: {min_value}, Max: {max_value}")
                        
                        modification_choice = input(
                            "Please select your next action:\n"
                            " 1. Reduce quantity of product (change the value of all F columns in the selected rows by a percentage)\n"
                            " 2. Reduce per-unit value of product (change the value of all I columns in the selected rows by a percentage)\n"
                            " 3. Reduce total value of product (change the value of all F and I columns in the selected rows by a percentage,)\nthis will also recalculate the J(αξια) value in all affected rows:\n "
                            " Selection(1-3): "
                        )
                        if modification_choice in ['1', '2', '3']:
                            modified_df = modify_values(df, min_value, max_value, int(modification_choice))
                            
                            # Get new file name
                            new_filename = input("Please select a name for your new file (just the file name, not including file type): ")
                            modified_df.to_excel(f"{new_filename}.xlsx", index=False)

                            # Copy the formatting from the original file to the new file
                            copy_formatting(filename, f"{new_filename}.xlsx")
                            
                            print(f"Your new file has been created: {new_filename}.xlsx with original formatting.")
                            
                            # Ask if the user wants to convert another file
                            another_file = input("Do you want to convert another file? (y/n): ").lower()
                            if another_file == 'n':
                                print("Exiting...")
                                return  # Exit the loop and program
                else:
                    print("Invalid file name, please try again.")
        elif action == '2':
            print("Exiting...")
            break
        else:
            print("Please select an applicable action (1, 2).")

if __name__ == "__main__":
    main()
