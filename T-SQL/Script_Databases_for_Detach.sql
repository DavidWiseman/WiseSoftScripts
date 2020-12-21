/* 
	Created By:		David Wiseman
	Website:		http://www.wisesoft.co.uk
	Description:
	Generates a script that can be used to detach all of the databases from a SQL Server instance.
	Run in SSMS and output results to text.
	                       
	** USE WITH CAUTION **
*/
SET NOCOUNT ON;
SELECT 'EXEC master.dbo.sp_detach_db @dbname = N' + QUOTENAME(name,'''') + ', @keepfulltextindexfile=N''true''
GO'
FROM sys.databases 
WHERE name NOT IN('master','msdb','tempdb','model');
