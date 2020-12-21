/* 
	Created By:		David Wiseman
	Website:		http://www.wisesoft.co.uk
	Description:
	Script to generate reports based on sys.dm_io_virtual_file_stats.
	See: http://msdn.microsoft.com/en-us/library/ms190326.aspx
*/
-- I/O Stats by Database
SELECT d.name as DatabaseName,
	ROUND(CAST(SUM(num_of_bytes_read+num_of_bytes_written) as float) / SUM(SUM(num_of_bytes_read+num_of_bytes_written)) OVER() *100,2) as [% Total I/O],
	ROUND(CAST(SUM(num_of_bytes_read) as float) / SUM(SUM(num_of_bytes_read)) OVER() *100,2) as [% Read I/O],
	ROUND(CAST(SUM(num_of_bytes_written) as float) / SUM(SUM(num_of_bytes_written)) OVER() *100,2) as [% Write I/O],
	ROUND(CAST(SUM(num_of_bytes_read+num_of_bytes_written)/(1024*1024*1024.0) as float),2) as [Total GB],	 
	ROUND(CAST(SUM(num_of_bytes_read)/(1024*1024*1024.0) as float),2) as [Read GB],
	ROUND(CAST(SUM(num_of_bytes_written)/(1024*1024*1024.0) as float),2) as [Write GB],
	SUM(io_stall) AS [I/O Total Wait ms],
	ISNULL(NULLIF(RIGHT(REPLICATE('0',
			--  Length of Max I/O stall in days over resultset (for dynamic padding - REPLICATE)
			LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall)) OVER() /86400000,0) AS VARCHAR),'')))
			--	I/O stall in days
			+ CAST(SUM(io_stall) / 86400000 AS VARCHAR) 
			--  Length of Max I/O stall in days over resultset (for dynamic padding - RIGHT)	
			,LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall)) OVER() /86400000,0) AS VARCHAR),'')))
			 + ' ',' '),'') 
			-- I/O stall formatted to HH:mm:ss	 
			 + LEFT(CONVERT(VARCHAR,DATEADD(s,SUM(io_stall)/1000,0),114),8) AS [I/O Total Wait Time {Days} HH:mm:ss],
	SUM(io_stall_read_ms) AS [I/O Read Wait ms]	,
	ISNULL(NULLIF(RIGHT(REPLICATE('0',
			--  Length of Max I/O read stall in days over resultset (for dynamic padding - REPLICATE)
			LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall_read_ms)) OVER() /86400000,0) AS VARCHAR),'')))
			--	I/O read stall in days
			+ CAST(SUM(io_stall_read_ms) / 86400000 AS VARCHAR) 
			--  Length of Max I/O read stall in days over resultset (for dynamic padding - RIGHT)	
			,LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall_read_ms)) OVER() /86400000,0) AS VARCHAR),'')))
			 + ' ',' '),'') 
			-- I/O read stall formatted to HH:mm:ss	 
			 + LEFT(CONVERT(VARCHAR,DATEADD(s,SUM(io_stall_read_ms)/1000,0),114),8) AS [I/O Read Wait Time {Days} HH:mm:ss]	,
	SUM(io_stall_write_ms) AS [I/O Write Wait ms],
	ISNULL(NULLIF(RIGHT(REPLICATE('0',
			--  Length of Max I/O write stall in days over resultset (for dynamic padding - REPLICATE)
			LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall_write_ms)) OVER() /86400000,0) AS VARCHAR),'')))
			--	I/O write stall in days
			+ CAST(SUM(io_stall_write_ms) / 86400000 AS VARCHAR) 
			--  Length of Max I/O write stall in days over resultset (for dynamic padding - RIGHT)	
			,LEN(ISNULL(CAST(NULLIF(MAX(SUM(io_stall_write_ms)) OVER() /86400000,0) AS VARCHAR),'')))
			 + ' ',' '),'') 
			-- I/O write stall formatted to HH:mm:ss	 
			 + LEFT(CONVERT(VARCHAR,DATEADD(s,SUM(io_stall_write_ms)/1000,0),114),8) AS [I/O Write Wait Time {Days} HH:mm:ss],	
	SUM(io_stall) / NULLIF(SUM(num_of_reads+num_of_writes),0)	 AS [Avg I/O Wait ms],
	SUM(io_stall_read_ms) / NULLIF(SUM(num_of_reads),0)	 AS [Avg Read I/O Wait ms], 
	SUM(io_stall_write_ms) / NULLIF(SUM(num_of_writes),0)	 AS [Avg Write I/O Wait ms],
	SUM(num_of_bytes_read+num_of_bytes_written)/NULLIF(SUM(num_of_reads+num_of_writes),0) AS [Avg I/O bytes],
	SUM(num_of_bytes_read)/NULLIF(SUM(num_of_reads),0) AS [Avg Read I/O bytes],
	SUM(num_of_bytes_written)/NULLIF(SUM(num_of_writes),0) AS [Avg Write I/O bytes],	
	CAST(MAX(sample_ms) / 86400000 AS VARCHAR) 
			-- I/O write stall formatted to HH:mm:ss	 
			 + ' ' + LEFT(CONVERT(VARCHAR,DATEADD(s,MAX(sample_ms)/1000,0),114),8) AS [Sample Time {Days} HH:mm:ss]  	 
FROM sys.dm_io_virtual_file_stats(null,null) vfs
JOIN sys.databases d ON vfs.database_id = d.database_id
GROUP BY d.name
ORDER BY [% Total I/O] DESC;
-- I/O Stats by file
SELECT d.name as DatabaseName,
	mf.name as logical_name,
	mf.physical_name,
	ROUND(CAST(SUM(num_of_byte