SET NOCOUNT ON			
SELECT '//NEED THIS NAMESPACE : using Data.DataImport;'

SELECT 
	CASE 
		WHEN t.[name] LIKE '%int' THEN 'record.SetInt32("'
		WHEN t.[name] = 'bit' THEN 'record.SetBoolean("'
		WHEN t.[name] = 'float' THEN 'record.SetDouble("'
		WHEN t.[name] IN ('decimal', 'money', 'numeric', 'smallmoney') THEN 'record.SetDecimal("'
		WHEN t.[name] = 'datetimeoffset' THEN 'record.SetDateTimeOffset("' 
		WHEN t.[name] IN ('binary', 'rowversion', 'varbinary') THEN 'record.SetDecimal("'
		WHEN t.[name] = 'uniqueidentifier' THEN 'record.SetGuid("'
		WHEN t.[name] IN ('date', 'time') THEN 'record.SetDateTime("'
		WHEN t.[name] LIKE '%datetime%' THEN 'record.SetDateTime("'
		WHEN t.[name] LIKE '%char' THEN 'record.SetString("'
		ELSE 'record.SetValue<' + UPPER(t.[name]) + '>("' 
	END + c.[name] + '", mapObject.' + c.[name] + ');' 
FROM sys.columns c 
INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
INNER JOIN sys.table_types tt ON tt.type_table_object_id = c.object_id
WHERE tt.user_type_id = TYPE_ID('dbo.Company_Import_tt')
ORDER BY c.column_id

