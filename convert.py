import pandas as pd
import mysql.connector

def excel_to_sql(excel_file, sql_file, table_name):
    # Read Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file)

    # Generate SQL statements
    sql_statements = []
    columns = ', '.join(df.columns)
    for index, row in df.iterrows():
        # Replace NaN values with None and handle other values
        values = [x if not pd.isna(x) else None for x in row]
        placeholders = ', '.join(['%s' for _ in values])
        sql_statement = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders});"
        sql_statements.append((sql_statement, tuple(values)))

    # Write SQL statements to a SQL file
    with open(sql_file, 'w', encoding='utf-8') as f:
        for statement, _ in sql_statements:
            f.write(statement + '\n')

    return sql_statements



if __name__ == "__main__":
    # Provide the path to your Excel file, desired SQL file, and table name
    excel_file_path = 'filepath.xlsx'
    sql_file_path = 'filepath.sql'
    # sql_file_path = 'investigations.sql'
    table_name = 'table_name'
    # table_name = 'vista_investigations'

    # Connect to MySQL database (replace with your database credentials)
    connection = mysql.connector.connect(
        host='host',
        user='user',
        password='',
        database='db_name'
    )

    # Execute SQL statements
    cursor = connection.cursor()
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} (id INT AUTO_INCREMENT PRIMARY KEY);")

    # Generate and execute parameterized SQL statements
    sql_statements = excel_to_sql(excel_file_path, sql_file_path, table_name)
    for statement, values in sql_statements:
        try:
            print("Executing SQL statement:")
            print(statement % values)
            cursor.execute(statement, values)
        except mysql.connector.Error as err:
            print(f"Error: {err}")

    connection.commit()

    # Close the connection
    cursor.close()
    connection.close()
