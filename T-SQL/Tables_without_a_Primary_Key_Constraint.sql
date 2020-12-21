SELECT s.name as SchemaName,o.name as ObjectName
FROM sys.objects o
JOIN sys.schemas s on o.[schema_id] = s.[schema_id]
WHERE o.type = 'U'
AND NOT EXISTS(SELECT 1
				FROM sys.indexes si 
				WHERE si.[object_id] = o.[object_id]
				AND is_primary_key = 1)
ORDER BY SchemaName,ObjectName