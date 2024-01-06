
from openpyxl import Workbook, load_workbook
import os
import re
import datetime


DB = "employee_assets.xlsx"


def open_workbook(file_name = "employee_assets.xlsx"):
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the directory of the script
    file_path = os.path.join(script_dir, file_name)  # Combine with the relative file path
    try:
        wb = load_workbook(file_path)
        return wb
    except FileNotFoundError:
        print(f"Error while opening: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error while opening: {e}")
        return None
    
def append_to_workbook(wb, data_to_append):

    try:
        ws = wb.active
        ws.append(data_to_append)  # Append a row with the provided data

        wb.save(DB)  # Save the changes
        print("Row appended and workbook closed successfully.")
        return True
    except Exception as e:
        print(f"Error while appending: {e}")
        return False
    
def close_workbook(wb):

    try:
        ws = wb.active
        wb.close()  # Close the workbook

        return True
    except Exception as e:
        print(f"Error while closing: {e}")
        return False
    
def save_data(data):
    
    confirm = input("Are you sure you want to append the following record ?\n"+str(data)+"\n(y/n)")
    
    while(confirm not in ["y","n"]):
        print("wrong choice please enter \'y\' or \'n\'")
        confirm = input("Are you sure you want to append the following record ?\n"+str(data)+"\n(y/n)")
        
    if confirm == "y":
            wb = open_workbook()
            result = append_to_workbook(wb, data)
            close_workbook(wb)
            return result
        
    elif confirm == "n":
        return False
    
def search_by_emp_id(emp_id_column, target_emp_id):

    try:
        wb = open_workbook()
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            current_emp_id = row[ord(emp_id_column) - ord('A')]  # Convert column letter to index
            if current_emp_id == target_emp_id:
                row_index = row[emp_id_column - 1]  # Assuming the emp_id is in the first column (adjust if needed)
                return (row_index, ord(emp_id_column) - ord('A') + 1)  # Return row and column indices

        # If emp_id is not found
        return (None, None)

    except Exception as e:
        print(f"Error: {e}")
        return (None, None)
    
def validate_14_digit(input_string):
    # Define a regex pattern for exactly 14 digits
    pattern = re.compile(r'^\d{14}$')

    # Use the pattern to match the input
    match = pattern.match(input_string)

    # If there is a match, the input is a 14-digit number
    return bool(match)

def id_validation(emp_id):
    while( not validate_14_digit(emp_id)):
        emp_id = str(input("please enter a valid id, should be 14 digit number: "))
    return emp_id


def add_employ():
    emp_name = input("please enter your name:  ")
    
    emp_id = str(input("please enter your id: "))

    emp_id = id_validation(emp_id)
    exist = search_by_emp_id(2, emp_id)
    if exist == (None, None):
        print("New Employee has no prior records, Please add your assets")
        result = add_asset(emp_id)
        save_data([emp_name] + result)

def add_asset(emp_id = "0"):
    
    if id_validation(emp_id) == emp_id: 
        asset = input("Asset type: \n1 - Laptop\n2 - HeadPhone\n3 - Mobile\n4 - Monitor\n")
        asset_sn = input("please enter asset serial number: ")
        note = input("note: ")
        timestamp = datetime.datetime.now()
        return [emp_id, asset, asset_sn, note, timestamp]
        




def main_menu():
    open_workbook()
    print("1 - add user")
    print("2 - add asset to user")
    print("3 - delete asset from user")
    print("4 - transfer asset")
    print("5 - list by:")
    x = int(input('Enter a number: '))
    
    if x == 1:
        add_employ()
    elif x == 2:
        add_asset()
if __name__ == "__main__":
    main_menu()
    # data = ["John Doe", "123456","Laptop", "1000", "2023-01-01"]
    # print(save_data(data))