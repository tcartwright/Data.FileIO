SELECT c.name,
	cast(COLUMNPROPERTY(c.object_id, c.name, 'IsIdentity') AS bit) AS [IsIdentity],
	t.name AS [TypeName],
	c.max_length,
	c.precision,
	c.scale,
	c.is_nullable
FROM sys.columns c 
INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
INNER JOIN sys.table_types tt ON tt.type_table_object_id = c.object_id
WHERE tt.user_type_id = TYPE_ID(@TableTypeName)
ORDER BY c.column_id