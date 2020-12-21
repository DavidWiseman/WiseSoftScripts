CREATE PROC [dbo].[BackupDatabases](
	-- Databases to backup in a comma-separated list.
	-- e.g. database1,database2,database3
	-- NULL = Backup all databases
	@Databases NVARCHAR(MAX)=NULL,
	-- Databases to be excluded from the backup in a comma-separated list
	-- e.g. database1,database2,database3
	-- NULL = Don't exclude any databases
	@ExcludedDatabases NVARCHAR(MAX)=NULL,
	-- Type of Backup to perform
	-- FULL = full database backup, DIFF = differential database Backup, TRAN = transaaction log backup
	@BackupType CHAR(4)='FULL',
	-- Backup Directory
	-- e.g. \\BackupServer\BackupShare or C:\Backups
	-- Multiple locations can be specified using a pipe symbol.
	-- e.g. C:\Backups|D:\Backups|E:\Backups
	-- Note: If you just want to split the backup into multiple parts, you can also specify the same location multiple times. e.g.
	-- C:\Backups|C:\Backups|C:\Backups
	-- When multiple locations are specifed, the file name is appended with the file number and number of files.  e.g. 1of3, 2of3, 3of3
	@BackupDir NVARCHAR(MAX),
	-- Run VERIFYONLY check after backup has completed
	-- 1 = Verify, 0 = Don't verify
	@Verify BIT=0,
	-- Perform backup with CHECKSUM option
	-- 1 = Perform CHECKSUM, 0 = Don't Perform CHECKSUM
	@CheckSum BIT=0,
	-- Option to perform DBCC CHECKDB command before backup
	-- 0 = don't perform DBCC check, 1 = perform DBCC check, 2 = perform DBCC check with physical_only option
	@PerformDBCC TINYINT=0,
	-- Option to remove backup files after a specified number of hours
	-- e.g. 24 = Keep backups for 1 day, 168 = Keep backups for 7 days (24*7)
	-- NULL = Don't remove backup files
	@RetainHours INT=NULL,
	-- Option to delete old backup files before performing backup. 
	-- Ideally you want to ensure that you have a valid backup before deleting old backup files so this option is best set to zero.
	-- 0 = remove after backup completed, 1 = remove before backup completed, NULL = Backup files not removed
	@DeleteBeforeBackup BIT=0,
	-- Option to debug this stored procedure
	-- 1 = Debug Mode (Print Commands), 0 = Execute Mode (Perform Backups)
	@Debug BIT=0
)
AS
/* 
	Created By:		David Wiseman
	Date:			2009-12-01
	Website:		http://www.wisesoft.co.uk
	Description:
	SQL Server database backup script.  Can be used to backup all databases on a SQL server instance or include/exclude specific databases.  
	Backups can be automatically deleted after a specified period of time.
	 
	Examples:
	EXEC dbo.BackupDatabases @BackupDir='C:\Backups'
	EXEC dbo.BackupDatabases @BackupDir='C:\Backups',@RetainHours=336 -- Delete after 2 weeks
	EXEC dbo.BackupDatabases @BackupDir='C:\Backups',@RetainHours=336, -- Delete after 2 weeks
							@ExcludedDatabases='nobackupdb1,nobackupdb2',
							@BackupType='DIFF'  							

	Requires SQL CLR IO Utility and dbo.SplitString function:
	http://www.wisesoft.co.uk/articles/sql_server_clr_io_utility.aspx 
	http://www.wisesoft.co.uk/scripts/t-sql_cte_split_string_function.aspx														

*/
SET NOCOUNT ON;
DECLARE @Database sysname;
DECLARE @FileNamePattern nvarchar(1024);
DECLARE @BackupCommand nvarchar(max);
DECLARE @BackupName nvarchar(max);
DECLARE @BackupDBs TABLE(name sysname);
DECLARE @ErrorMessage NVARCHAR(4000);
DECLARE @ErrorSeverity INT;
DECLARE @ErrorState INT;
DECLARE @ErrorCount INT;
DECLARE @BackupLocations NVARCHAR(MAX)
SET @ErrorCount = 0

-- Check that a valid backup type is specified
IF @BackupType NOT IN('FULL','DIFF','TRAN')
BEGIN;
	RAISERROR ('Invalid Backup Type Specified. Options: FULL,DIFF,TRAN',11,1);
	RETURN;
END;

IF @Databases IS NULL
BEGIN -- Backup of all databases required (excluding databases that are not applicable for backup type)
	INSERT INTO @BackupDBs(name)
	SELECT name
	FROM sys.databases db 
	WHERE source_database_id IS NULL -- Exclude database snapshots
		AND name <> 'tempdb' -- Exclude tempdb database
		AND [state] = 0 --ONLINE databases only
		AND is_in_standby=0 -- Exclude databases in standby mode
		AND NOT (@BackupType='DIFF' AND name='master') -- Exclu