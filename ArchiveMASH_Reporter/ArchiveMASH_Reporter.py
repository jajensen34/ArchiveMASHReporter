import pyodbc
import pandas as pd
import os
from pathlib import Path
import datetime
import sys


def main():
    print("************************************************")
    print("  ArchiveMUSH V1.0 - The ArchiveMASH Assistant")
    print("************************************************")

    # Gather SQL Server and DB Name used by ArchiveMASH (Assume we're using the current Windows login creds).
    while True:
        try:
            server_name = input("Enter the SQL Server name: ")
            DBName = input("Enter the name of the ArchiveMASH database: ")
            # server_name = 'evdasql2016'
            # DBName = 'archivem'
            # Connect to the SQL server
            conn = pyodbc.connect(
                'Driver={SQL Server};'
                'Server=' + server_name + ';'
                'Database=' + DBName + ';'
                'Trusted_Connection=yes;'
            )
                #'Driver={ODBC Driver 17 for SQL Server};Server=' + server_name + ';Database=' + DBName + ';Trusted_Connection=yes;')
        except:
            print('Connection Unsuccessful.', sys.exc_info()[0])
            continue

        else:
            print(f'Connected to {server_name}.')
            break

    # Provide list of reports and ask for selection.
    main_menu = ['Migration Status Report', 'Migration Rate Report', 'Change Migration Status']
    user_id = None
    group_id = None
    user_name = ""

    selection = None
    while not selection:
        selection = get_menu_option(main_menu, "#########Main Menu#########")

    if selection == 'Migration Rate Report':
        group_selection = None
        group_query = 'SELECT GroupName from Users GROUP BY GroupName'
        group_list = run_simple_query(group_query, conn)
        while not group_selection:
            group_selection = get_menu_option(group_list, "Specify Migration Group:")
            group_id = group_selection[0]
        yes_no = ''
        while yes_no.lower() not in ('y', 'n'):
            try:
                yes_no = input('Do you want to run report for a single user? (Y/N): ')
            except yes_no.lower() not in ('y', 'n'):
                print('Invalid Entry.  Try again.')
                continue

        if yes_no.lower() == 'y':
            user_selection = None
            user_query = "SELECT UserId, DisplayName FROM users where GroupName = '" + group_id + "'"
            user_list = run_simple_query(user_query, conn)
            while not user_selection:
                user_selection = get_menu_option(user_list, "Select User: ")
                user_id = user_selection[0]
                user_name = user_selection[1]

    # Run selected report.
    sql_run = get_report(selection)
    run_report(sql_run, selection, conn, group_id, user_id, user_name)


def run_report(report_file, name, conn, group_id, user_id,
               user_name):  # Function loads and runs report from SQL query file.
    query_text = load_query(report_file)
    query_change = ""
    # If there's a userID or GroupID limiter in the query, add those.
    if group_id != None:
        for line in query_text:
            stripped_line = line.strip()
            new_line = stripped_line
            if stripped_line.find("<GroupID>") != -1:
                new_line = stripped_line.replace("<GroupID>", group_id)
            elif (stripped_line.find("u.UserID > 0") != -1 and user_id != None):
                user_text = 'u.UserID = ' + str(user_id)
                new_line = stripped_line.replace("u.UserID > 0", user_text)
            query_change += new_line + "\n"
        query_text = query_change
        query_change = ""
        stripped_line = ""
        new_line = ""

    # take the text pulled from the query file and run it.
    with conn.cursor():
        report_query = ''.join(query_text)
    # print(report_query)
    # input()
    if group_id != None:
        content = "Group: " + group_id + "    "
        if user_name != "":
            content += "User: " + user_name + "\n"
    content = pd.read_sql_query(report_query, conn)

    print()
    print(content)
    # Print report to file.
    header_text = write_header(name, group_id, user_name)
    write_report(name, content, header_text)


# Show completion status and present location of report.


def create_report(file_name):  # function creates a report file and directory.
    file_name = file_name + datetime.datetime.now().strftime("-%m-%d-%Y_%H%M%S.csv")
    parent_dir = os.path.dirname(__file__)
    path = os.path.join(parent_dir, "Reports")
    if not os.path.isdir(path):
        os.mkdir(path)
    abs_file = os.path.join(path, file_name)
    return abs_file


def write_header(report_name, group_id, user_name):
    header_text = report_name + "\n \n"
    if (report_name == 'Migration Rate Report'):
        header_text += "Group: " + group_id
        if user_name != "":
            header_text += " ,User: " + user_name + "\n \n"
        else:
            header_text += "\n \n"
        header_text += "Start Time,  Stop Time,  Minutes Elapsed \n"
    elif (report_name == 'Migration Status Report'):
        header_text += "Group Name, Archive Name, Archive ID, Status, # Migrated, # Not Migrated, # Failed, MB Migrated \n"
    return header_text


def write_report(name, content, header_text):  # function writes data to the report file.
    filename = create_report(name)
    content.to_csv(filename, index=False, header=True)
    with open(filename) as lines:
        line = lines.readlines()

    line[0] = header_text

    with open(filename, "w") as lines:
        lines.writelines(line)


def load_query(file_name):  # function loads the specified SQL query to be run.
    abs_file = os.path.abspath(os.path.join("lib", file_name))
    query = open(abs_file, 'r')
    return query


def run_simple_query(query, conn):
    output = pd.read_sql_query(query, conn)
    menu_list = output.values.tolist()
    # print(menu_list)
    # stop = input('')
    # for row in cursor:
    # menu_list.append(row)
    return menu_list


def get_menu_option(mlist, mname):
    print(mname)
    for index, opt in enumerate(mlist, start=1):
        print(f"{index}:  {opt}")
    try:
        user_input: int = int(input("Selection: "))
        user_input -= 1
        print(mlist[user_input])
        if user_input < 0 or user_input >= len(mlist):
            print(f"{user_input} is an invalid entry.  Try again.")
            return None
    except ValueError as ve:
        print(f"Could not convert {ve} to an integer.  Try again.")
        return None

    return mlist[user_input]


def get_report(report_name):
    reports = {
        "Migration Status Report": "MigrationStatus.sql",
        "Migration Rate Report": "MigPerformance.sql"
    }
    sql_run = reports.get(report_name)
    return sql_run


if __name__ == '__main__':
    main()
