SET NOCOUNT ON

DECLARE @tablename sysname = 'dbo.Company_Import',
	@sql varchar(max)

SELECT @sql = 'IF TYPE_ID(''' + @tablename + '_tt'') IS NOT NULL BEGIN
	DROP TYPE ' + @tablename + '_tt
END
GO

CREATE TYPE ' + @tablename + '_tt AS TABLE(
'

SELECT @sql += STUFF((
	SELECT CHAR(9) + ',' + QUOTENAME([c].name) 
			+ ' ' + t.name 
			+ CASE 
				WHEN t.name = 'datetime2' THEN '(' + CAST(c.precision AS VARCHAR(20)) + ')' 
				WHEN t.name = 'decimal' THEN '(' + CAST(c.precision AS VARCHAR(20)) + ',' + CAST(c.scale AS VARCHAR(20)) + ')' 
				WHEN t.name LIKE '%char' THEN '(' + CASE WHEN [c].max_length = -1 THEN 'MAX' ELSE CAST([c].max_length AS VARCHAR(20)) END  + ')' 
				ELSE '' 
			END 
			+ CASE WHEN c.is_nullable = 1 THEN ' NOT NULL' ELSE ' NULL' END + CHAR(10)
	FROM sys.columns c 
	INNER JOIN sys.types t ON t.system_type_id = c.system_type_id
	WHERE c.object_id = OBJECT_ID(@tablename)
	FOR XML PATH('')), 1, 2, '')


SELECT @sql += '
	,[ImportErrors] varchar(MAX) NOT NULL
);'; 

select @sql
