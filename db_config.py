import mysql.connector
from mysql.connector import Error

def get_db_connection():


    connection = None
    try:
        connection = mysql.connector.connect(
            host='localhost',
            user='admin_user',
            password='admin_password',
            database='excel_to_database'
        )
        if connection.is_connected():
            print("Successfully connected to the database")
    except Error as e:
        print(f"Error: '{e}'")

    return connection