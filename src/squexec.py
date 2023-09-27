import pyodbc
import openpyxl
import datetime


def connect_to_server(server, database, username, password):
    try:
        connection_string = (
            "DRIVER=SQL Server;"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password}"
        )
        
        connection = pyodbc.connect(connection_string)
        return connection  # Return the connection object
    except Exception as e:
        print(f"Error: {str(e)}")
        return None  # Return None in case of an error


def save_res_to_excel(result, excel_output_file, cursor):
    workbook = openpyxl.Workbook()
    sheet = workbook.active  # Use the active sheet

    # Print headers based on the column names
    if cursor.description is not None:
        headers = [i[0] for i in cursor.description]
        print(headers)
        sheet.append(headers)

    for row in result:
        rowlist = list(row)
        sheet.append(rowlist)

    workbook.save(excel_output_file)


def execute_query(query, server, database, username, password):
    connection = connect_to_server(server, database, username, password)

    if connection:
        cursor = connection.cursor()
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        connection.close()
        return result
    else:
        return None


def main():
    print("#################################################################\n")
    print("\t\tQuery Executer\n")
    print("#################################################################\n")

    server = input("Enter Server Address:\t\t")
    database = input("Enter Database Name:\t\t")
    username = input("Enter Username:\t\t\t")
    password = input("Enter Password:\t\t\t")

    # Query input
    query_lines = []
    print("Input your query. You can input multiline queries.")
    print("When you press enter on an empty line, the query will be complete\n")
    print("Query (press Enter on an empty line to finish):\n")

    while True:
        uinput = input()
        if uinput == "":
            break
        else:
            query_lines.append(uinput)

    query = "\n".join(query_lines)

    print(f"Your query:\n{query}")

    confirm = input("Do you want to proceed with the query? (y/n):\t\t")

    if confirm.lower() == "y":
        result = execute_query(query, server, database, username, password)
        if result is not None:
            ts = datetime.datetime.now().strftime("%Y-%m-%d_%H.%M.%S")
            excel_output_file = f"output_{ts}.xlsx"
            connection = connect_to_server(server, database, username, password)
            cursor = connection.cursor()
            save_res_to_excel(result, excel_output_file, cursor)
            print(f"Query results saved to {excel_output_file}")
        else:
            print("Failed to connect to the server or execute the query.")
    elif confirm.lower() == "n":
        print("You entered 'n'. No query will be executed.")
    else:
        print("Invalid choice")


if __name__ == "__main__":
    main()
