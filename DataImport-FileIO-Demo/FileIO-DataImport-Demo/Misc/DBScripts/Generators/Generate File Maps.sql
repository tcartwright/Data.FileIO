

SET NOCOUNT ON			

SELECT 
	CASE 
		WHEN t.name LIKE '%int' THEN 'this.' + c.name + ' = validator.GetRowValue<int>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name = 'bit' THEN 'this.' + c.name + ' = validator.GetRowValue<bool>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name = 'float' THEN 'this.' + c.name + ' = validator.GetRowValue<Double>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name IN ('decimal', 'money', 'numeric', 'smallmoney') THEN 'this.' + c.name + ' = validator.GetRowValue<Decimal>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name = 'datetimeoffset' THEN 'this.' + c.name + ' = validator.GetRowValue<DateTimeOffset>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name IN ('binary', 'rowversion', 'varbinary') THEN 'this.' + c.name + ' = validator.GetRowValue<byte[]>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name = 'uniqueidentifier' THEN 'this.' + c.name + ' = validator.GetRowValue<Guid>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name IN ('date', 'time') THEN 'this.' + c.name + ' = validator.GetRowValue<DateTime>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name LIKE '%datetime%' THEN 'this.' + c.name + ' = validator.GetRowValue<DateTime>(row, "' + c.name + '", ref errors' + fn.[nullable]
		WHEN t.name LIKE '%char' THEN 'this.' + c.name + ' = row.' + c.name + ';'
		ELSE 'this.' + c.name + ' = validator.GetRowValue<' + t.name + '>(row, "' + c.name + '", ref errors' + fn.[nullable]
	END
FROM sys.columns c 
INNER JOIN sys.types t ON t.system_type_id = c.system_type_id
CROSS APPLY(SELECT nullable = CASE WHEN c.[IS_NULLABLE] = 1 THEN ', isNullable: true);' ELSE ');' END) fn
WHERE c.object_id = OBJECT_ID('Company_Import')
ORDER BY [c].column_id
