CREATE PROC dbo.spCreateSnapshot( 
	@Database sysname, -- Source database name 
	@SnapshotName sysname, -- Snapshot database name 
	@Folder NVARCHAR(MAX), -- Location to store snapshot files 
	@Debug BIT=0 -- Set to 1 to print SQL statement, 0 to create snapshot 
)
AS
/* 
	Created By David Wiseman, 06/03/2009
	http://www.wisesoft.co.uk
	Creates a database snapshot of a given database.
	
	USAGE:
	exec spCreateSnapshot @database='AdventureWorks',@SnapshotName='AdventureWorks_SS',@folder='C:\',@debug=0
*/
/*	Updated By David Wiseman, 30/04/2012
	sys.master_files is now used to support database mirrors
*/
DECLARE @SQL NVARCHAR(MAX);
IF RIGHT(@Folder,1) <> '\'
	SET @Folder = @Folder + '\';

SELECT @SQL = 'CREATE DATABASE ' + QUOTENAME(@SnapshotName) + ' ON
' +
STUFF((SELECT ',(NAME = ' + QUOTENAME(Name) + ',FILENAME=' + QUOTENAME(@Folder + @SnapshotName + '_' + Name + '.ss','''') + ')
'
FROM sys.master_files
WHERE database_id = DB_ID(@Database)
AND [type] = 0 --ROWS
FOR XML PATH(''),TYPE).value('.','NVARCHAR(MAX)'),1,1,' ')
+ ' AS SNAPSHOT OF ' + QUOTENAME(@Database);

IF @Debug = 1
	PRINT @SQL;
ELSE 
	EXEC sp_executesql @SQL;