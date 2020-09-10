import pyodbc
import pandas as pd
import os
from pathlib import Path
import datetime
import tkinter as tk

def main():
    #create the window.
    w_main = tk.Tk()
    w_main.title("ArchiveMUSH - ArchiveMASH Helper")
    #title = tk.Label(text="ArchiveMUSH - ArchiveMASH Helper")
    #title.pack()
    f_main = tk.Frame(master=w_main, width=800, height=400)
    f_main.pack()
    w_main.mainloop()

    #Gather SQL Server and DB Name used by ArchiveMASH (Assume we're using the current Windows login creds).
    server_name = input("Enter the SQL Server name: ")
    DBName = input("Enter the name of the ArchiveMASH database: ")


    #Connect to the SQL server
    conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};Server='+server_name+';Database='+DBName+';Trusted_Connection=yes;')

    #Provide list of reports and ask for selection.

    #Run selected report.
    run_report('MigrationStatus.sql', 'Migration_Status', conn)


def run_report(report_file, name, conn):    #Function loads and runs report from SQL query file.
    query_text = load_query(report_file)
   #create query from file, removing leading whitespace.
    with conn.cursor() as cursor:
       report_query = ' '.join(query_text)

    content = pd.read_sql_query(report_query,conn)
    print()
    print(content)
    # Print report to file.
    write_report(name, content)

#Show completion status and present location of report.


def create_report(file_name): #function creates a report file and directory.
    file_name = file_name + datetime.datetime.now().strftime("-%m-%d-%Y_%H%M%S.csv")
    directory = os.path.dirname(__file__)
    p = Path(directory + '/Reports')
    p.mkdir(exist_ok=True)
    abs_file = os.path.join(p, file_name)
    return abs_file

def write_report(name, content): #function writes data to the report file.
    filename = create_report(name)
    content.to_csv(filename, index = False, header=True)

def load_query(file_name): #function loads the specified SQL query to be run.
    directory = os.path.dirname(__file__)
    abs_file = os.path.join(directory, file_name)
    query = open(file_name, 'r')
    return query




if __name__ == '__main__':
    main()