using Data.DataImport;
using Data.DataImport.Interfaces;
using Data.DataImport.SqlRecordImporters;
using FileIODemo.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace FileIODemo.DataImport
{
    public class EnumerablesDataImport
	{
		public static void EnumerableImport1(IEnumerable<Company> data)
		{
			Console.WriteLine("Running EnumerableImport1");

            var insertParameters = new ImportEnumerableParameters<Company>(data, "[dbo].[ImportCompanies]")
            {
                ProcedureParameterName = "@data",
                CustomSqlMapper = Common.MapValues,
                ErrorColumnName = "ImportErrors"
            };

            var repo = new DataImportRepository();
			using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString))
			{
				conn.Open();

				//pulling the sqlmetadata manually, can be re-used if re-importing this same data multiple times
				var sqlMetaData = repo.GetSqlMetadata(null, conn, "[dbo].[Company_Import_tt]");
				insertParameters.SqlMetadata.AddRange(sqlMetaData);

				using (var tran = conn.BeginTransaction())
				{
					repo.ImportEnumerable<Company>(tran, insertParameters);

					tran.Commit();
				}
			}
		}


		public static void EnumerableImport2(IEnumerable<Company> data)
		{
			Console.WriteLine("Running EnumerableImport2");

			List<FakeObject> extraList = new List<FakeObject>();

			IImportEnumerableParameters<Company> insertParameters = new ImportEnumerableParameters<Company>(data, "[dbo].[ImportCompanies]");
			insertParameters.ProcedureParameterName = "@data";
			//insertParameters.CustomSqlMapper = Common.MapValues;
			insertParameters.ErrorColumnName = "ImportErrors";

			var repo = new DataImportRepository();
			using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString))
			{
				conn.Open();

				//pulling the sqlmetadata manually, can be re-used if re-importing this same data multiple times
				var sqlMetaData = repo.GetSqlMetadata(null, conn, "[dbo].[Company_Import_tt]");
				insertParameters.SqlMetadata.AddRange(sqlMetaData);

				#region this part shows how to include an EXTRA table parameter / enumerable.
				//IMPORTANT NOTE: It is important that you DO NOT attach the extra table parameter if the list is empty as an empty list will cause an exception
				if (extraList.Any())
				{
					//NOTE: This extra parameter example would require a second table type(Fake_Data_tt), which actually does not exist in my example

					//first create the insert parameters
					IImportEnumerableParameters<FakeObject> fakeDataInsertParameters = new ImportEnumerableParameters<FakeObject>(extraList, repo.GetSqlMetadata(null, conn, "[dbo].[Fake_Data_tt]"));
					//create the enumerablesqlrecord, the will be our parameters "value"
					var fakeData = new EnumerableSqlRecord<FakeObject>(fakeDataInsertParameters);

					insertParameters.ExtraParameters.Add(new SqlParameter("@fakeData", SqlDbType.Structured) { Value = fakeData });
				}
				#endregion this part shows how to include an EXTRA table / enumerable.

				using (SqlTransaction tran = conn.BeginTransaction())
				{
					repo.ImportEnumerable<Company>(tran, insertParameters);

					tran.Commit();
				}
			}
		}

	}
}
