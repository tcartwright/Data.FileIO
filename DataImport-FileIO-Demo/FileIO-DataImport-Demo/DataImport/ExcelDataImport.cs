using Data.FileIO.DataImport;
using Data.FileIO.DataImport.Interfaces;
using Data.FileIO.Interfaces;
using Data.FileIO.Parsers;
using FileIODemo.Data;
using System;
using System.Configuration;
using System.Data.SqlClient;

namespace FileIODemo.DataImport
{
    static class ExcelDataImport
	{
		public static void ExcelImport1(string path)
		{
			Console.WriteLine("Running ExcelImport1");

			/*This example shows importing a Company class from a file to a database using Streaming TVP.  */

			//same excel file parser for reading
			IExcelFileParser parser = new ExcelFileParser(path, "Companies", 2, 3);
			IImportFileParameters<Company> insertParameters = new ImportFileParameters<Company>(parser, "dbo.ImportCompanies", "Company_Import_tt");
			/*
			 * REQUIRED Parameters:
			 *		parser: The file parser
			 *		procedureName: The stored proc to execute the import
			 *		tableTypeName: The name of the table type that the parer is passed in to OR
			 *		sqlMetaData: If, the table type name is not passed in, this must be. It is a list of the sql metadata columns that represent the table type. The FileDataImportRepository.GetSqlMetadata can build these for you.
			 * 
			 * OPTIONAL Properties:
			 *		CustomValidator: Provides custom validation to validate the object, allowing for custom errors 
			 *		Data: This represents the data rows from the parser object passed into the ctor, and should not be touched.
			 *		ErrorColumnName: This is the column that the errors will be written to with an xml representation of the file import errors. Will only be automatically be written to when no mapping is provided.
			 *		ExtraParameters: These are extra parameters that can be passed into the stored procedure.
			 *		FileValuesMapper: Provides a custom mapper for the file parser to map the file row to the object
			 *		Parser: This represents the parser itself that was passed into the ctor, and should not be touched
			 *		ProcedureName: The name of the import stored procedure
			 *		ProcedureParameterName: The named of the table type parameter in the stored proc
			 *		SqlMetadata: the sqlmeta data that represents the columns of the table type
			 *		TableTypeName: The name of the table type
			 *				NOTE: Either the TableTypeName OR the SqlMetadata must be provided.
			 *		Validator: The validator to use for file to object mapping	
			 *		CustomSqlMapper: Provides custom mapping from the class being imported to the sqlrecord which contains the sqlmetadata 
			 */
			insertParameters.ProcedureParameterName = "@data";
			insertParameters.CustomSqlMapper = Common.MapValues;
			insertParameters.ErrorColumnName = "ImportErrors";

			var repo = new FileDataImportRepository();

			using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString))
			{
				conn.Open();

				using (var tran = conn.BeginTransaction())
				{
					repo.ImportFile<Company>(tran, insertParameters);

					tran.Commit();
				}
			}
		}

	}
}
