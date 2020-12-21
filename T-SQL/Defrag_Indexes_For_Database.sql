/* Prerequsites
	Requires dbo.fnSplitString function:
	http://www.wisesoft.co.uk/scripts/t-sql_split_string_function_while_loop.aspx
	or
	http://www.wisesoft.co.uk/scripts/t-sql_cte_split_string_function.aspx
*/
-- Remove objects from DB if they already exists (uncomment if required)
/*
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[index_maintenance_log]') AND type in (N'U'))
DROP TABLE [dbo].[index_maintenance_log]
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[index_maintenance_batch]') AND type in (N'U'))
DROP TABLE [dbo].[index_maintenance_batch]
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DefragIndexesForDB]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[DefragIndexesForDB]
GO
*/
CREATE TABLE dbo.index_maintenance_batch(
	batch_id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_index_maintenance_batch PRIMARY KEY(batch_id),
	start_time DATETIME NOT NULL,
	end_time DATETIME NULL,
	reorganize_threshold TINYINT NOT NULL,
	rebuild_threshold TINYINT NOT NULL,
	database_name SYSNAME NOT NULL,
	is_online BIT NOT NULL,
	[sort_in_tempdb] BIT NOT NULL,
	page_count_threshold INT NOT NULL,
	excluded_tables NVARCHAR(MAX)NULL
)
GO
CREATE TABLE dbo.index_maintenance_log(
	log_id INT IDENTITY(1,1) CONSTRAINT PK_index_maintenance_log PRIMARY KEY(log_id),
	batch_id INT NOT NULL CONSTRAINT FK_index_maintenance_log_index_maintenance_batch FOREIGN KEY REFERENCES dbo.index_maintenance_batch(batch_id),
	[schema_name] SYSNAME NOT NULL,
	[object_name] SYSNAME NOT NULL,
	index_name SYSNAME NOT NULL,
	index_type_desc NVARCHAR(60) NOT NULL,
	partition_number INT NOT NULL,
	avg_fragmentation_in_percent float NOT NULL,
	is_rebuild bit NOT NULL,
	is_online bit NOT NULL,
	start_time DATETIME NOT NULL,
	end_time DATETIME NULL
)
GO
CREATE PROC [dbo].[DefragIndexesForDB](
	-- Threshold to perform index maintenance. Fragmentation levels below this value will be ignored.
	@ReorganizeThreshold TINYINT=15,
	-- Threshold to rebuild indexes rather than reorganize. 
	-- If you don't want to use rebuild, set the value to >100.  
	-- If you want to rebuild rather than reorganize, set the value to the same as the ReorganizeThreshold
	@RebuildThreshold TINYINT=30,
	-- Database to defrag
	@DatabaseName SYSNAME,
	-- If specified all rebuilds will be done online.  In cases where that is not possible, the index will be reorganized, regardless of the RebuildThreshold
	-- The online option is only available in enterprise, developer and evaluation editions of SQL Server.  Set the rebuild threhold greater than 100 to use a reorganize instead.
	-- Note: It's strongly recommended to perform index maintenance out of hours, even with the online option set to 1
	@Online BIT=1,
	-- If specified, index rebuild statements will be printed and won't be run
	@Debug BIT=0,
	-- Option to sort index in tempdb
	@SortInTempDB BIT=1,
	-- Used to exclude small indexes.  
	@PageCountThreshold INT=512, -- Default value = 512 pages/4MB
	-- Option to exclude tables from index maintenance
	-- Should be a comma separated string of tables names to exclude
	-- e.g. 'dbo.MyExcludedTable,dbo.MyExcludedTable2,dbo.MyExcludedTable3'
	@ExcludeTables NVARCHAR(MAX)=NULL
)
/*  Created:	13/10/2009
	Updated:	01/03/2011
	Version:	1.02
	Author:		David Wiseman
	Website:	http://www.wisesoft.co.uk
	Purpose:	Procedure to defrag indexes for a given database, based on levels of fragmentation specified.
	Notes:		Ignores small indexes (less than 512 pages/4MB) and disabled indexes.
				Requires dbo.fnSplitString function:
				http://www.wisesoft.co.uk/scripts/t-sql_split_string_function_while_loop.aspx
				or
				http://www.wisesoft.co.uk/scripts/t-sql_cte_split_string_function.aspx
	Example:
	
	/*		Reorganize all indexed in "AdventureWorks" database with an avg_fragmentation_in_percent between 15 and <30.
			Rebuild indexes in "AdventureWorks" data