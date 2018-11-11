using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Data.FileIO.Core;
using Data.DataImport.Readers;
using Data.FileIO.Interfaces;
using Data.FileIOTests.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Data.FileIO.Writers.Tests
{
	[TestClass()]
	public class CsvWriterTests
	{
		private string _destinationPath;
		private static int _employeesCount = 5000;

		public TestContext TestContext { get; set; }

		#region test init and cleanup
		[TestInitialize]
		public void TestInit()
		{
			do
			{
				//loop until we find a file name that does not exist for our destination, should only be one time
				_destinationPath = Path.Combine(Path.GetTempPath(), "CsvWriterTests", String.Format("{0}_{1}.csv", TestContext.TestName, Path.GetFileNameWithoutExtension(Path.GetTempFileName())));
			} while (File.Exists(_destinationPath));
		}

		[TestCleanup]
		public void TestCleanup()
		{
			if (File.Exists(_destinationPath))
			{
				//this code is not valid for unit test, just here for debugging
				if (Debugger.IsAttached)
				{
					var process = new ProcessStartInfo("Excel.exe", _destinationPath);
					var p = Process.Start(process);
				}
				else
				{
					File.Delete(_destinationPath);
				}
			}
			_destinationPath = string.Empty;
		}
		#endregion test init and cleanup

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvWriterTests")]
		public void WriteCSVTestWithTypeNoMaps()
		{
			CsvWriter writer = new CsvWriter();
			writer.WriteData(_destinationPath, Employee.GetTestEmployees(_employeesCount), '"');

			Assert.IsTrue(File.Exists(_destinationPath));
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvWriterTests")]
		public void WriteCSVTestWithTypeNoMapsNoQuote()
		{
			CsvWriter writer = new CsvWriter();
			writer.WriteData(_destinationPath, Employee.GetTestEmployees(_employeesCount));

			Assert.IsTrue(File.Exists(_destinationPath));
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvWriterTests")]
		public void WriteCSVTestWithTypeNoMapsNoQuoteNoHeader()
		{
			CsvWriter writer = new CsvWriter();
			writer.WriteData(_destinationPath, Employee.GetTestEmployees(_employeesCount), writeHeaders: false);

			Assert.IsTrue(File.Exists(_destinationPath));
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvWriterTests")]
		public void WriteCSVTestWithReader()
		{
			CsvWriter writer = new CsvWriter();
			writer.WriteData(_destinationPath, Employee.GetTestEmployees(_employeesCount).GetDataReader());

			Assert.IsTrue(File.Exists(_destinationPath));
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvWriterTests")]
		public void WriteCSVTestWithTypeWithMapsNoQuote()
		{
			CsvColumnInfoList<Employee> columnInfoList = new CsvColumnInfoList<Employee>();
			//IList<ICsvColumnInfo<Employee>> columnInfoList = new List<ICsvColumnInfo<Employee>>();
			columnInfoList.Add("SSN", obj => obj.L2SSN);
			columnInfoList.Add("Name", obj => obj.L1Name, '"'); //set a quoted identifier on this field only
			columnInfoList.Add("Hire Date", obj => obj.HireDate.ToShortDateString());
			columnInfoList.Add("Hourly Rate", obj => obj.HourlyRate);

			CsvWriter writer = new CsvWriter();
			writer.WriteData(_destinationPath, Employee.GetTestEmployees(_employeesCount), columnInfoList);

			Assert.IsTrue(File.Exists(_destinationPath));
		}
	}
}
