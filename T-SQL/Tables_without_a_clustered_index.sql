SELECT s.name AS SchemaName,
		o.name  AS ObjectName
FROM sys.indexes si
JOIN sys.objects o ON si.object_id = o.object_id
JOIN sys.schemas s ON o.schema_id = s.schema_id
WHERE si.type=0 --HEAP
AND o.type ='U' -- USER_TABLE