import pandas as pd
import os
import re

def sanitize_column_name(name):
    # Replace invalid characters with underscores
    return re.sub(r'\W|^(?=\d)', '_', name)

def preprocess_dataframe(df):
    # Drop rows with all NaN values
    df.dropna(how='all', inplace=True)
    
    # Drop columns with all NaN values
    df.dropna(axis=1, how='all', inplace=True)
    
    # Explicitly cast columns to string type to avoid dtype issues
    df = df.astype(str)
    
    # Fill NaN values with empty strings
    df.fillna('', inplace=True)
    
    # Drop columns with any NaN values or blank spaces
    df = df.loc[:, (df != '').all(axis=0)]
    
    # Drop rows with any NaN values or blank spaces
    df = df.loc[(df != '').all(axis=1)]
    
    # Drop columns with any blank spaces
    df = df.loc[:, (df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) != '').all(axis=0)]
    
    # Drop rows with any blank spaces
    df = df.loc[(df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) != '').all(axis=1)]
    
    # Sanitize column names
    df.columns = [sanitize_column_name(col) for col in df.columns]
    
    return df

def convert_excel_to_db(file_path, db_connection):
    # Determine the file extension
    file_extension = os.path.splitext(file_path)[1]

    # Read the file based on its extension
    if file_extension == '.xlsx' or file_extension == '.xls':
        df = pd.read_excel(file_path, header=1)  # Assuming headers are in the second row
    elif file_extension == '.csv':
        df = pd.read_csv(file_path, header=8)
    else:
        raise ValueError("Unsupported file format. Please provide an Excel or CSV file.")

    # Preprocess the DataFrame
    df = preprocess_dataframe(df)

    # Convert DataFrame to records
    records = df.to_dict(orient='records')

    cursor = db_connection.cursor()

    # Drop the existing table if it exists
    cursor.execute("DROP TABLE IF EXISTS data_table")

    # Create table based on DataFrame columns
    columns = ', '.join([f"`{col}` VARCHAR(255)" for col in df.columns])
    create_table_sql = f"CREATE TABLE data_table ({columns})"
    cursor.execute(create_table_sql)

    # Insert data into the table
    for record in records:
        # Prepare the SQL insert statement
        columns = ', '.join([f"`{key}`" for key in record.keys()])
        placeholders = ', '.join(['%s'] * len(record))
        sql = f"INSERT INTO data_table ({columns}) VALUES ({placeholders})"
        
        # Execute the insert statement
        cursor.execute(sql, tuple(record.values()))

    # Commit the transaction
    db_connection.commit()
    cursor.close()