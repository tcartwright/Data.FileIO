using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Data.DataImport;
using Data.DataImport.Interfaces;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Data.FileIO.DataImport;
using Data.FileIO.DataImport.Interfaces;
using Data.FileIO.Interfaces;
using Data.FileIO.Parsers;
using Data.FileIOTests.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Data.FileIO.Repository.Tests
{
	//http://blogs.interknowlogy.com/2014/08/29/unit-testing-using-localdb/
	[TestClass()]
	public class DataImportRepositoryTests
	{

		public TestContext TestContext { get; set; }

		[TestInitialize]
		public void Initialize()
		{
			//use a different data directory for each test, so the db will be clean for each test
			string directory = Path.Combine(this.TestContext.TestDeploymentDir, this.TestContext.TestName, string.Empty);
			AppDomain.CurrentDomain.SetData("DataDirectory", directory);
		}


		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_DataImportRepositoryTests")]
		[DeploymentItem(@"UnitTestData.mdf", "ImportExcelFileTest")] //deploy to a test specific folder, the folder MUST match the test name to match the DataDirectory
		[DeploymentItem(@"Files\1095_Import_WithData.xlsx", "Files")]
		public void ImportExcelFileTest()
		{

			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);
			IImportFileParameters<Employee> insertParameters = new ImportFileParameters<Employee>(parser, "dbo.Import_Employees", "EmployeeData");
			//insertParameters.Mapper = this.MapValues;	

			IFileDataImportRepository repo = new FileDataImportRepository();

			using (SqlConnection conn = new SqlConnection(GetConnectionString()))
			{
				conn.Open();
				DeleteEmployees(conn); //just in case they switched the connection string to a regular sql server

				Stopwatch sw = Stopwatch.StartNew();
				repo.ImportFile<Employee>(conn, insertParameters);
				sw.Stop();
				Debug.WriteLine("Import elapsed time: {0}", sw.Elapsed);

				int employeesCount = GetEmployeesCount(conn);

				Assert.IsTrue(employeesCount > 1); //as the number of the rows may change just assert we got any data added to the table
			}
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_DataImportRepositoryTests")]
		[DeploymentItem(@"UnitTestData.mdf", "ImportEmptyExcelFileTest")] //deploy to a test specific folder, the folder MUST match the test name to match the DataDirectory
		[DeploymentItem(@"Files\1095_Import_Template.xlsx", "Files")]
		public void ImportEmptyExcelFileTest()
		{

			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_Template.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);
			IImportFileParameters<Employee> insertParameters = new ImportFileParameters<Employee>(parser, "dbo.Import_Employees", "EmployeeData");
			//insertParameters.Mapper = this.MapValues;	

			var repo = new FileDataImportRepository();
			
			using (SqlConnection conn = new SqlConnection(GetConnectionString()))
			{
				conn.Open();
				DeleteEmployees(conn); //just in case they switched the connection string to a regular sql server

				Stopwatch sw = Stopwatch.StartNew();
				repo.ImportFile<Employee>(conn, insertParameters);
				sw.Stop();
				Debug.WriteLine("Import elapsed time: {0}", sw.Elapsed);

				int employeesCount = GetEmployeesCount(conn);

				Assert.AreEqual(0, employeesCount); //as the number of the rows may change just assert we got any data added to the table
			}
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_DataImportRepositoryTests")]
		[DeploymentItem(@"UnitTestData.mdf", "ImportEnumerableTest")] //deploy to a test specific folder, the folder MUST match the test name to match the DataDirectory
		public void ImportEnumerableTest()
		{
			int listCount = 5000;
			IEnumerable<Employee> data = Employee.GetTestEmployees(listCount);

			IImportEnumerableParameters<Employee> insertParameters = new ImportEnumerableParameters<Employee>(data, "dbo.Import_Employees");
			insertParameters.CustomSqlMapper = CustomEmployeesMapSqlRecord;

			var repo = new DataImportRepository();
			using (SqlConnection conn = new SqlConnection(GetConnectionString()))
			{
				conn.Open();
				DeleteEmployees(conn); //just in case they switched the connection string to a regular sql server

				insertParameters.SqlMetadata.AddRange(repo.GetSqlMetadata(null, conn, "EmployeeData"));

				Stopwatch sw = Stopwatch.StartNew();
				repo.ImportEnumerable<Employee>(conn, insertParameters);
				sw.Stop();
				Debug.WriteLine("Import elapsed time: {0}", sw.Elapsed);

				int employeesCount = GetEmployeesCount(conn);

				Assert.AreEqual(listCount, employeesCount);
			}
		}


		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_DataImportRepositoryTests")]
		[DeploymentItem(@"UnitTestData.mdf", "ImportEmptyEnumerableTest")] //deploy to a test specific folder, the folder MUST match the test name to match the DataDirectory
		public void ImportEmptyEnumerableTest()
		{
			int listCount = 0;
			IEnumerable<Employee> data = null;

			IImportEnumerableParameters<Employee> insertParameters = new ImportEnumerableParameters<Employee>(data, "dbo.Import_Employees");
			insertParameters.CustomSqlMapper = CustomEmployeesMapSqlRecord;

			IDataImportRepository repo = new DataImportRepository();
			using (SqlConnection conn = new SqlConnection(GetConnectionString()))
			{
				conn.Open();
				DeleteEmployees(conn); //just in case they switched the connection string to a regular sql server

				insertParameters.SqlMetadata.AddRange(repo.GetSqlMetadata(null, conn, "EmployeeData"));

				Stopwatch sw = Stopwatch.StartNew();
				repo.ImportEnumerable<Employee>(conn, insertParameters);
				sw.Stop();
				Debug.WriteLine("Import elapsed time: {0}", sw.Elapsed);

				int employeesCount = GetEmployeesCount(conn);

				Assert.AreEqual(listCount, employeesCount);
			}
		}

		#region private helper methods

		public void CustomEmployeesMapSqlRecord(Employee mapObject, Microsoft.SqlServer.Server.SqlDataRecord record, int rowIndex, IEnumerable<string> errors)
		{
			record.SetValue(record.GetOrdinal("SSN"), mapObject.L2SSN);
			record.SetValue(record.GetOrdinal("Name"), mapObject.L1Name);
			string errXml = errors.Count() > 0 ? FileIOHelpers.ErrorsToXml(errors, rowIndex) : null;
			record.SetValue(record.GetOrdinal("Errors"), errXml);
		}

		private string GetConnectionString()
		{
			//have to vary the catalog so this error does not happen when using localdb: Cannot attach file {file} as database 'UnitTests' because this database name is already attached with file {other file} 
			SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["UnitTests"].ConnectionString);
			if (builder.DataSource.ToLower().Contains("(localdb)")) //only rewrite the catalog if using local db
			{
				builder.InitialCatalog = this.TestContext.TestName;
			}
			return builder.ConnectionString;
		}

		private void DeleteEmployees(SqlConnection conn)
		{
			using (var command = conn.CreateCommand())
			{
				command.CommandText = "Delete from Employees";
				command.CommandType = System.Data.CommandType.Text;
				command.ExecuteNonQuery();
			}
		}

		private int GetEmployeesCount(SqlConnection conn)
		{
			using (var command = conn.CreateCommand())
			{
				command.CommandText = "Select count(*) from Employees";
				command.CommandType = System.Data.CommandType.Text;
				return (int)command.ExecuteScalar();
			}
		}

		//example map values for employee
		private void MapValues(ref Employee mapObject, int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors)
		{
			mapObject.L2SSN = row.L2SSN;
			mapObject.L1Name = row.L1Name;
		}

		#endregion	private helper methods
	}
}
