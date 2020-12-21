/* 
	Created By:		David Wiseman
	Website:		http://www.wisesoft.co.uk
	Description:
	Script databases for attach.  Run this script before detaching your databases to generate an attach script for each database.
	Also checks db ownership chaining, trustworthy and broker enabled settings which are normally lost as part of the detach/attach process.
	Run in SSMS and output results to text.
*/
SET NOCOUNT ON;
DECLARE @Database SYSNAME;
DECLARE @IsTrustworthy BIT;
DECLARE @IsBrokerEnabled BIT;
DECLARE @IsDBChaining BIT;
DECLARE @SQL NVARCHAR(MAX);
-- Cursor to loop over each DB
DECLARE cDatabases CURSOR FAST_FORWARD
	FOR SELECT name,is_trustworthy_on,is_broker_enabled,is_db_chaining_on
	FROM sys.databases 
	WHERE name NOT IN('master','msdb','tempdb','model')
	AND state=0; -- ONLINE
	
OPEN cDatabases;

WHILE 1=1
BEGIN;
	FETCH NEXT FROM cDatabases INTO @Database,@IsTrustworthy,@IsBrokerEnabled,@IsDBChaining;
	IF @@FETCH_STATUS <> 0
		BREAK;
	-- Generate script to attach database.
	SET @SQL = 'USE ' + QUOTENAME(@Database) + '
	SELECT ''CREATE DATABASE '' + QUOTENAME(DB_NAME()) + '' ON'' +
	STUFF((SELECT '',
	(FILENAME = '' + QUOTENAME(physical_name,'''''''') + '')''
	FROM sys.database_files
	FOR XML PATH(''''),TYPE).value(''.'',''NVARCHAR(MAX)''),1,1,'''') + ''
	FOR ATTACH
GO''';
	
	exec sp_executesql @SQL;
	
	-- Ensure trustworthy, db chaining and broker enabled options are maintained after the attach.
	IF @IsTrustworthy=1
	BEGIN;
		SELECT 'ALTER DATABASE ' + QUOTENAME(@Database) + ' SET TRUSTWORTHY ON
GO';
	END;
	IF @IsBrokerEnabled=1
	BEGIN;
		SELECT 'ALTER DATABASE ' + QUOTENAME(@Database) + ' SET ENABLE_BROKER
GO';
	END;
	IF @IsDBChaining=1
	BEGIN;
		SELECT 'ALTER DATABASE ' + QUOTENAME(@Database) ' SET DB_CHAINING ON
GO';
	END;
END;
	
CLOSE cDatabases;
DEALLOCATE cDatabases;