CREATE FUNCTION fnSplitString(@str nvarchar(max),@sep nvarchar(max))
RETURNS @tbl TABLE(value nvarchar(max))
AS
BEGIN
	DECLARE @idx1 INT;
	DECLARE @idx2 INT;
	SET @idx1=0;
	WHILE @idx1 >-1
	BEGIN;
		SELECT @idx2 =  CHARINDEX(@sep,@str,@idx1);
		IF @idx2 > 0
		BEGIN;
			INSERT INTO @tbl(value)
			SELECT SUBSTRING(@str,@idx1,@idx2-@idx1)
			SET @idx1 = @idx2+1;
		END;
		ELSE
		BEGIN;
			INSERT INTO @tbl(value)
			SELECT SUBSTRING(@str,@idx1,LEN(@str)+1-@idx1)
			SET @idx1 = -1;
		END;
	END;
	RETURN;
END;