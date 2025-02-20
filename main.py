from db_config import get_db_connection
from excel_to_db import convert_excel_to_db
import os

def main():
    excel_file_path = os.path.abspath('files/exercise 2.csv') 
    connection = get_db_connection()
    
    try:
        convert_excel_to_db(excel_file_path, connection)
        print("Data has been successfully imported into the database.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        connection.close()

if __name__ == "__main__":
    main()