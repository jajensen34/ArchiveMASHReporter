import pyodbc


#Gather SQL Server and DB Name used by ArchiveMASH (Assume we're using the current Windows login creds).
server_name = input("Enter the SQL Server name: ")
DBName = input("Enter the name of the ArchiveMASH database: ")
user_name = input("Enter user name: ")
user_pw = input("Enter password: ")

#Connect to the SQL server
conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server={server_name};'
                      'Database={DBName};'
                      'UID={user_name};'
                      'PWD={user_pw}')

cursor = conn.cursor()

#Provide list of reports and ask for selection.


#Run selected report.


#Show completion status and present location of report.
