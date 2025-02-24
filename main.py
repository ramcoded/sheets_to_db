from db_config import get_db_connection
from excel_to_db import convert_excel_to_db
import os

def get_user_input():
    print("Enter the values of columns and rows to skip from the left and top of the file.")
    skip_columns_input = input("Enter the number of columns to skip (e.g., 3 to skip columns [0, 1, 2]): ")
    skip_rows_input = input("Enter the number of rows to skip (e.g., 3 to skip rows [0, 1, 2]): ")

    if skip_columns_input:
        skip_columns = list(range(int(skip_columns_input)))
    else:
        skip_columns = []

    if skip_rows_input:
        skip_rows = list(range(int(skip_rows_input)))
    else:
        skip_rows = []

    return skip_columns, skip_rows

def main():
    excel_file_path = os.path.abspath('files/exercise2.xlsx')
    connection = get_db_connection()
    
    try:
        skip_columns, skip_rows = get_user_input()
        print(f"Skipping columns: {skip_columns}")
        print(f"Skipping rows: {skip_rows}")
        convert_excel_to_db(excel_file_path, connection, skip_columns=skip_columns, skip_rows=skip_rows)
        print("Data has been successfully imported into the database.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        connection.close()

if __name__ == "__main__":
    main()