import re
import sql_metadata
import pyodbc


def mssql_connect():
    
    return cnxn


def sql_query(query):
    server = 'mow03-sql56c'
    database = 'agro3'
    # username = 'your_username'
    # password = 'your_password'
    driver= '{ODBC Driver 13 for SQL Server}'
    try:
        cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE=master;Trusted_Connection=yes')
    except Exception, e:
        print Server
        print type(e), e
        print '_________'

    cursor = cnxn.cursor()
    cursor.execute(query)
   
    result =[]
    row = cursor.fetchone()
    while row:
        # print (str(row[0]) + " " + str(row[1]))
        print row
        result.append(row)
        row = cursor.fetchone()
        
    return result
    



if __name__ == '__main__':
  

  mssql_connect
#   a=sql_query("select @@version")
a=sql_query("""call{sp_BlitzCache}""")
  
print a
  
    # SQL="""
    # INSERT INTO #tt290 (_Q_000_F_000RRef, _Q_000_F_001_TYPE, _Q_000_F_001_RTRef, _Q_000_F_001_RRRef, _Q_000_F_002RRef, _Q_000_F_003RRef, _Q_000_F_004RRef, _Q_000_F_005RRef, _Q_000_F_006RRef, _Q_000_F_007, _Q_000_F_008, _Q_000_F_009RRef) SELECT T1.Fld34219RRef, T1.Fld34223_TYPE, T1.Fld34223_RTRef, T1.Fld34223_RRRef, T1.Fld34220RRef, T1.Fld34221RRef, T1.Fld34222RRef, T1.Fld34226RRef, T1.Fld34225RRef, T1.Fld34228Balance_, T1.Fld34229Balance_, T1.Fld34224RRef FROM (SELECT T2.Fld34224RRef AS Fld34224RRef, T2.Fld34223_TYPE AS Fld34223_TYPE, T2.Fld34223_RTRef AS Fld34223_RTRef, T2.Fld34223_RRRef AS Fld34223_RRRef, T2.Fld34222RRef AS Fld34222RRef, T2.Fld34221RRef AS Fld34221RRef, T2.Fld34220RRef AS Fld34220RRef, T2.Fld34219RRef AS Fld34219RRef, T2.Fld34225RRef AS Fld34225RRef, T2.Fld34226RRef AS Fld34226RRef, CAST(SUM(T2.Fld34229Balance_) AS NUMERIC(38, 8)) AS Fld34229Balance_, CAST(SUM(T2.Fld34228Balance_) AS NUMERIC(38, 8)) AS Fld34228Balance_ FROM (SELECT T3._Fld34224RRef AS Fld34224RRef, T3._Fld34223_TYPE AS Fld34223_TYPE, T3._Fld34223_RTRef AS Fld34223_RTRef, T3._Fld34223_RRRef AS Fld34223_RRRef, T3._Fld34222RRef AS Fld34222RRef, T3._Fld34221RRef AS Fld34221RRef, T3._Fld34220RRef AS Fld34220RRef, T3._Fld34219RRef AS Fld34219RRef, T3._Fld34225RRef AS Fld34225RRef, T3._Fld34226RRef AS Fld34226RRef, CAST(SUM(T3._Fld34229) AS NUMERIC(33, 8)) AS Fld34229Balance_, CAST(SUM(T3._Fld34228) AS NUMERIC(32, 8)) AS Fld34228Balance_ FROM _AccumRgT34238 T3 WHERE T3._Period = @P1 AND ((((T3._Fld34227RRef = @P2) AND EXISTS(SELECT 1 FROM _InfoRg31000 T4 WHERE ((T4._RecorderTRef = @P3 AND T4._RecorderRRef = @P4)) AND (T3._Fld34219RRef = T4._Fld31046RRef) AND (T3._Fld34221RRef = T4._Fld31092RRef))) AND (T3._Fld34220RRef IN (SELECT T5._Fld31069RRef AS Q_002_F_000RRef FROM _InfoRg31000 T5 WHERE (T5._RecorderTRef = @P5 AND T5._RecorderRRef = @P6)) OR (T3._Fld34220RRef = @P7)))) GROUP BY T3._Fld34224RRef, T3._Fld34223_TYPE, T3._Fld34223_RTRef, T3._Fld34223_RRRef, T3._Fld34222RRef, T3._Fld34221RRef, T3._Fld34220RRef, T3._Fld34219RRef, T3._Fld34225RRef, T3._Fld34226RRef HAVING (CAST(SUM(T3._Fld34229) AS NUMERIC(33, 8))) <> @P8 OR (CAST(SUM(T3._Fld34228) AS NUMERIC(32, 8))) <> @P9 UNION ALL SELECT T6._Fld34224RRef AS Fld34224RRef, T6._Fld34223_TYPE AS Fld34223_TYPE, T6._Fld34223_RTRef AS Fld34223_RTRef, T6._Fld34223_RRRef AS Fld34223_RRRef, T6._Fld34222RRef AS Fld34222RRef, T6._Fld34221RRef AS Fld34221RRef, T6._Fld34220RRef AS Fld34220RRef, T6._Fld34219RRef AS Fld34219RRef, T6._Fld34225RRef AS Fld34225RRef, T6._Fld34226RRef AS Fld34226RRef, CAST(CAST(SUM(CASE WHEN T6._RecordKind = 0.0 THEN -T6._Fld34229 ELSE T6._Fld34229 END) AS NUMERIC(27, 8)) AS NUMERIC(27, 2)) AS Fld34229Balance_, CAST(CAST(SUM(CASE WHEN T6._RecordKind = 0.0 THEN -T6._Fld34228 ELSE T6._Fld34228 END) AS NUMERIC(26, 8)) AS NUMERIC(27, 3)) AS Fld34228Balance_ FROM _AccumRg34218 T6 WHERE (T6._Period > @P10 OR T6._Period = @P11 AND (T6._RecorderTRef > @P12 OR T6._RecorderTRef = @P13 AND T6._RecorderRRef >= @P14)) AND T6._Period < @P15 AND T6._Active = 0x01 AND ((((T6._Fld34227RRef = @P16) AND EXISTS(SELECT 1 FROM _InfoRg31000 T7 WHERE ((T7._RecorderTRef = @P17 AND T7._RecorderRRef = @P18)) AND (T6._Fld34219RRef = T7._Fld31046RRef) AND (T6._Fld34221RRef = T7._Fld31092RRef))) AND (T6._Fld34220RRef IN (SELECT T8._Fld31069RRef AS Q_002_F_000RRef FROM _InfoRg31000 T8 WHERE (T8._RecorderTRef = @P19 AND T8._RecorderRRef = @P20)) OR (T6._Fld34220RRef = @P21)))) GROUP BY T6._Fld34224RRef, T6._Fld34223_TYPE, T6._Fld34223_RTRef, T6._Fld34223_RRRef, T6._Fld34222RRef, T6._Fld34221RRef, T6._Fld34220RRef, T6._Fld34219RRef, T6._Fld34225RRef, T6._Fld34226RRef HAVING (CAST(CAST(SUM(CASE WHEN T6._RecordKind = @P22 THEN -T6._Fld34229 ELSE T6._Fld34229 END) AS NUMERIC(27, 8)) AS NUMERIC(27, 2))) <> @P23 OR (CAST(CAST(SUM(CASE WHEN T6._RecordKind = @P24 THEN -T6._Fld34228 ELSE T6._Fld34228 END) AS NUMERIC(26, 8)) AS NUMERIC(27, 3))) <> @P25) T2 GROUP BY T2.Fld34224RRef, T2.Fld34223_TYPE, T2.Fld34223_RTRef, T2.Fld34223_RRRef, T2.Fld34222RRef, T2.Fld34221RRef, T2.Fld34220RRef, T2.Fld34219RRef, T2.Fld34225RRef, T2.Fld34226RRef HAVING (CAST(SUM(T2.Fld34229Balance_) AS NUMERIC(38, 8))) <> @P26 OR (CAST(SUM(T2.Fld34228Balance_) AS NUMERIC(38, 8))) <> @P27) T1
    # """
    # tables = sql_metadata.get_query_tables(SQL)
    # print("Parsing SQL Query for tables...")

    # Table_List=[]


    # for table in tables:
    #     if table[0] =='_':
    #         Table_List.append(table)

    # print(Table_List)

