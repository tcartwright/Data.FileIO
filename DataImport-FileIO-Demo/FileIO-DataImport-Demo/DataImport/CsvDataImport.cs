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
    static class CsvDataImport
	{
		public static void CsvImport1(string path)
		{
			Console.WriteLine("Running CsvImport1");
			ICsvFileParser parser = new CsvFileParser(path);
			
			IImportFileParameters<Company> insertParameters = new ImportFileParameters<Company>(parser, "dbo.ImportCompanies", "Company_Import_tt");
			insertParameters.ProcedureParameterName = "@data";
			insertParameters.CustomSqlMapper = Common.MapValues;
			insertParameters.ErrorColumnName = "ImportErrors";

			var repo = new FileDataImportRepository();

			using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString))
			{
				conn.Open();

				using (SqlTransaction tran = conn.BeginTransaction())
				{
					repo.ImportFile<Company>(tran, insertParameters);

					tran.Commit();
				}
			}
		}

	}
}
