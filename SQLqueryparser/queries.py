class sql_queries():
    def __init__(self):


        global SQL_get_indexes
        SQL_get_indexes = """
        SELECT DB_NAME() as DatabaseName, OBJECT_NAME(ind.OBJECT_ID) AS TableName,
            ind.name AS IndexName, indexstats.index_type_desc AS IndexType,
            indexstats.avg_fragmentation_in_percent
        FROM sys.dm_db_index_physical_stats(DB_ID('ITIL_DB'), null, NULL, NULL, NULL) indexstats
            --FROM sys.dm_db_index_physical_stats(DB_ID('ITIL_DB'), OBJECT_ID('ITIL_SYNC_INTERFACE'), NULL, NULL, NULL) indexstats 
            INNER JOIN sys.indexes ind
            ON ind.object_id = indexstats.object_id
                AND ind.index_id = indexstats.index_id
        ORDER BY indexstats.avg_fragmentation_in_percent DESC 
        """

        global SQL_TOP10_queries
        SQL_TOP10_queries = """
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
