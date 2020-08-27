import pyodbc
import pandas as pd


#Gather SQL Server and DB Name used by ArchiveMASH (Assume we're using the current Windows login creds).
server_name = input("Enter the SQL Server name: ")
DBName = input("Enter the name of the ArchiveMASH database: ")


#Connect to the SQL server
conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};Server='+server_name+';Database='+DBName+';Trusted_Connection=yes;')



#Provide list of reports and ask for selection.


#Run selected report.
with conn.cursor() as cursor:
    migrate_status = ' '.join((
               " select u.GroupName,u.DisplayName,u.SourceID,",
                "case when statusid in (1,2,3,4,5,8,9) then 'InProgress' when statusid in (6,7) then 'Complete' else 'Queued' end as [Status],",
                "SUM(Migrated) as [Migrated],",
                "SUM(NotMigrated) as [NotMigrated],",
                "SUM(Failed) as [Failed],",
                "SUM(MigratedMB) as [MigratedMB] from users u",
                "inner join",
                "(",
                    "select TargetID,ISNULL(SUM([0]),0) as [NotMigrated],",
                    "ISNULL(SUM([1]),0) as [Migrated],",
                    "ISNULL(SUM([2]),0) + ISNULL(SUM([4]),0) as [Failed],",
                    "ISNULL(SUM([3]),0) as [Skipped],",
                    "SUM([Migrated])/1024/1024 as [MigratedMB]",
                "from",
                    "(",
                        "select targetid,m.statusid as [StatusCount],count(*) as [MigratedCount],",
                        "case when m.statusid = 1 then 'Migrated' else 'NotMigrated' end as [StatusSize],",
                        "sum(cast(messagesize as bigint)) as [SizeB]",
                        "from messages m with (NOLOCK)",
                        "inner join folders f on f.folderid = m.folderid",
                        "inner join users u on u.userid = f.userid",
                        "group by GroupName,targetid,m.statusid",
                    ") as tbl",
                    "Pivot",
                    "(",
                        "SUM(MigratedCount) for StatusCount in ([0],[1],[2],[3],[4])",
                    ") as PivotTable",
                    "Pivot",
                    "(",
                        "SUM(SizeB) for StatusSize in ([Migrated])",
                    ") as PivotTable2",
                    "group by targetid",
                ") p",
                "on p.TargetID = u.targetid",
                "group by u.GroupName,u.DisplayName,u.SourceId,u.statusid",
                "order by u.GroupName,u.DisplayName",
    ))
    rpt_status = pd.read_sql_query(migrate_status,conn)
    print(rpt_status)

#Show completion status and present location of report.
