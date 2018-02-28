# import re
# import sql_metadata
import pyodbc
import pandas 
#import queries
from lxml import etree
from termcolor import colored


def sql_query(query, database):
    server = 'mow03-opt06'
    # database = 'agro3'
    # username = 'your_username'
    # password = 'your_password'
    driver= '{ODBC Driver 13 for SQL Server}'
    # print colored('Connectinng to {}', 'green').format(server)
    try:
        
        cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';Trusted_Connection=yes;autocommit=True')
    except Exception, e:
        print server
        print type(e), e
              
    DF=pandas.read_sql_query(query, cnxn)
    #print(DF)
    # print(type(DF))
    cnxn.close()
    return DF
    
def sql_query1(query):
    server = 'mow03-sql56c'
    # database = 'agro3'
    # username = 'your_username'
    # password = 'your_password'
    driver= '{SQL Server Native Client 11.0}'
    # print colored('Connectinng to {}', 'green').format(server)
    try:
        
        cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE=master;Trusted_Connection=yes;autocommit=True')
    except Exception, e:
        print server
        print type(e), e
    
    cursor = cnxn.cursor()         
    cursor.execute(query)
    
    while cursor.nextset():   # NB: This always skips the first resultset
        try:
            for row in cursor.fetchall():
                print row
            
            break
        except pyodbc.ProgrammingError:
            continue
    
    for row in cursor.fetchall():
        print row
    return 


if __name__ == '__main__':
    global SQL_get_indexes
    SQL_get_indexes = """
        SET NOCOUNT ON;
        --USE {0}
        --GO
        SELECT OBJECT_NAME(ind.OBJECT_ID) AS TableName,
            ind.name AS IndexName, indexstats.index_type_desc AS IndexType,
            round(indexstats.avg_fragmentation_in_percent, 0, -1) as [%frag]
            --FROM sys.dm_db_index_physical_stats(DB_ID('{0}'), null, NULL, NULL, NULL) indexstats
            FROM sys.dm_db_index_physical_stats(DB_ID('{0}'), OBJECT_ID('{1}'), NULL, NULL, NULL) indexstats 
            INNER JOIN sys.indexes ind
            ON ind.object_id = indexstats.object_id
                AND ind.index_id = indexstats.index_id
        ORDER BY indexstats.avg_fragmentation_in_percent DESC 
        """

    global SQL_get_indexes11
    SQL_get_indexes11 = """
        SET NOCOUNT ON;
        
        SELECT 
            DB_NAME() as DB,
            OBJECT_NAME(ind.OBJECT_ID) AS TableName,
            ind.name AS IndexName, indexstats.index_type_desc AS IndexType,
            round(indexstats.avg_fragmentation_in_percent, 0, -1) as [%frag]
            --FROM sys.dm_db_index_physical_stats(DB_ID('{0}'), null, NULL, NULL, NULL) indexstats
            FROM sys.dm_db_index_physical_stats(DB_ID('{0}'), OBJECT_ID('{1}'), NULL, NULL, NULL) indexstats 
            INNER JOIN sys.indexes ind
            ON ind.object_id = indexstats.object_id
                AND ind.index_id = indexstats.index_id
        ORDER BY indexstats.avg_fragmentation_in_percent DESC 
        """
    global SQL_get_tableinfo 
    SQL_get_tableinfo = """
SET NOCOUNT ON;

SELECT
    DB_NAME() as DB,
    t.Name AS TableName,
    p.rows AS Rows,
    CAST(ROUND((SUM(a.used_pages) / 128.00), 2) AS NUMERIC(36, 2)) AS Used_MB
FROM sys.tables t
    INNER JOIN sys.indexes i ON t.OBJECT_ID = i.object_id
    INNER JOIN sys.partitions p ON i.object_id = p.OBJECT_ID AND i.index_id = p.index_id
    INNER JOIN sys.allocation_units a ON p.partition_id = a.container_id
WHERE t.name = '{1}'
GROUP BY t.Name, p.Rows
    """

    global SQL_TOP10_queries
    SQL_TOP10_queries = """
        SET NOCOUNT ON;
        SELECT TOP 10

            qs.execution_count,
            qs.total_logical_reads, qs.last_logical_reads,
            qs.total_logical_writes, qs.last_logical_writes,
            qs.total_worker_time,
            qs.last_worker_time,
            qs.total_elapsed_time/1000000 total_elapsed_time_in_S,
            qs.last_elapsed_time/1000000 last_elapsed_time_in_S,
            qs.last_execution_time,
            qp.query_plan
        FROM sys.dm_exec_query_stats qs
            CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) qt
            CROSS APPLY sys.dm_exec_query_plan(qs.plan_handle) qp
        -- ORDER BY qs.total_logical_reads DESC -- logical reads
        -- ORDER BY qs.total_logical_writes DESC -- logical writes
        -- ORDER BY qs.total_worker_time DESC -- CPU time
        ORDER BY total_elapsed_time_in_S DESC 
            """

    

  
#   mssql_connect
#   a=sql_query("select @@version")
    print colored('Select TOP 10 statemens', 'green')
    a=sql_query(SQL_TOP10_queries, 'master')
    print colored(a,'red')
    qp = []
    qp = a['query_plan'].tolist()
    # print(qp[0])
    Table_List=[]
    d = {}
    i=0
    for q in qp:
        if q == None: continue
        i=i+1
        
        print colored('Parsing {} QueryPlan', 'green').format(i)
        #Parse XML Query Plan for DB and Tables in it into Table+List
        xml = etree.fromstring(q)
        # print xml.get('Table')
        for x in xml.iter():
            # fff=x.get()
            Tab = str(x.get('Table')).replace("[", "").replace("]", "")
            
            DB = str(x.get('Database')).replace("[", "").replace("]", "")
           
            if Tab != 'None' and DB != 'tempdb' and DB != 'None' and DB != 'NaN':
                d = dict(db =DB, table=Tab)
                Table_List.append(d)
    
        #Conver to Pandas Dataframe
    df = pandas.DataFrame(Table_List)
        #Dedupe DF
    df.drop_duplicates(inplace=True)
    print( df)
     
    allinfo=pandas.DataFrame()

    htm=open("c:\www\index.html", "w")
    JoinedDF = pandas.DataFrame()

    for index, row in df.iterrows():
        
        # print '111'
        # print SQL_get_tableinfo.format(row['db'], row['table'])
        # print row['db'], row['table']
        tabinfo=sql_query(SQL_get_tableinfo.format(row['db'], row['table']), row['db'])   
        if not tabinfo.empty:
            indexinfo=sql_query(SQL_get_indexes.format(row['db'], row['table']), row['db'])
            print colored(tabinfo.to_string(index=False), 'yellow')
            print colored(indexinfo.to_string(index=False), 'magenta')
            #print (tabinfo.merge(indexinfo, left_on='TableName', right_on='TableName', how='outer' ))
            MergedDF = tabinfo.merge(indexinfo, left_on='TableName', right_on='TableName', how='outer' )
            JoinedDF = JoinedDF.append( MergedDF, ignore_index = True)
            #JoinedDF = JoinedDF.append( indexinfo.to_string(index=False), ignore_index = True )

    JoinedDF.sort_values(by = ['DB','IndexType', '%frag'], kind= 'mergesort',  inplace = True, ascending=False )
    htm.write(JoinedDF.to_html(index = False))    
    #print JoinedDF
    htm.close()
    print colored('END!', 'red')
#all_info=tabinfo.join(indexinfo, on='TableName', how='outer' )
#print (all_info)





