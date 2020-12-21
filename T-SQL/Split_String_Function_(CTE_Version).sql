CREATE FUNCTION fnSplitString(@str nvarchar(max),@sep nvarchar(max))
RETURNS TABLE
AS
RETURN
	WITH a AS(
		SELECT CAST(0 AS BIGINT) as idx1,CHARINDEX(@sep,@str) idx2
		UNION ALL
		SELECT idx2+1,CHARINDEX(@sep,@str,idx2+1)
		FROM a
		WHERE idx2>0
	)
	SELECT SUBSTRING(@str,idx1,COALESCE(NULLIF(idx2,0),LEN(@str)+1)-idx1) as value
	FROM a
