using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Data.DataImport.Readers;
using Data.FileIO.Core;
using Data.FileIO.Interfaces;
using Data.FileIO.Writers;
using Data.FileIOTests.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Data.FileIO.Writers.Tests
{
	[TestClass()]
	[DeploymentItem("Files\\1095_Import_Template.xlsx", "Files")]
	public class ExcelWriterTests
	{
		private string _destinationPath;
		private string _sourcePath;
		private static int _employeesCount = 50;

		public TestContext TestContext { get; set; }

		#region test init and cleanup
		[TestInitialize]
		public void TestInit()
		{
			_sourcePath = Path.Combine(this.TestContext.DeploymentDirectory, "Files\\1095_Import_Template.xlsx");
			do
			{
				//loop until we find a file name that does not exist for our destination, should only be one time
				_destinationPath = Path.Combine(Path.GetTempPath(), "ExcelWriterTests", String.Format("{0}_{1}.xlsx", TestContext.TestName, Path.GetFileNameWithoutExtension(Path.GetTempFileName())));
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
					ValidateSpreadSheet(_destinationPath);
					var process = new ProcessStartInfo("Excel.exe", _destinationPath);
					var p = Process.Start(process);
				}
				else
				{
					File.Delete(_destinationPath);
				}
			}
			_sourcePath = string.Empty;
			_destinationPath = string.Empty;
		}
		#endregion test init and cleanup

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelWriterTests")]
		public void WriteFileWithNoTemplate()
		{
			int id = 1;
			var employees = Employee.GetTestEmployees(_employeesCount).ToList();

			ColumnInfoList<Employee> employeesMap = new ColumnInfoList<Employee>();
			//IList<IColumnInfo<Employee>> employeesMap = new List<IColumnInfo<Employee>>();
			employeesMap.Add("A", "Id", (obj) => IncrementId(ref id), cellType: ExcelCellType.Number);
			employeesMap.Add("B", "Name", (obj) => obj.L1Name);
			employeesMap.Add("C", "SSN", (obj) => obj.L2SSN);

			ColumnInfoList<Employer> employerMap = new ColumnInfoList<Employer>();
			//IList<IColumnInfo<Employer>> employerMap = new List<IColumnInfo<Employer>>();
			employerMap.Add("A", "Name", (obj) => obj.L7Name);
			employerMap.Add("B", "City", (obj) => obj.L11City);
			employerMap.Add("C", "State", (obj) => obj.L12State);
			employerMap.Add("D", "Zip", (obj) => obj.L13Zip, cellType: ExcelCellType.Number);

			using (ExcelWriter writer = new ExcelWriter(1, 2))
			{
				//completely control sheet creation by turning off header generation and passing in custom column infos
				writer.GenerateHeadersFromType = false;
				writer.WriteDataToSheet("Employees", employees, employeesMap);
				writer.WriteDataToSheet("Employer", Employer.GetEmployer(), employerMap);

				writer.GenerateHeadersFromType = true;
				writer.WriteDataToSheet("Employees No Maps", employees);

				writer.WriteTo(_destinationPath);
			}
			Assert.IsTrue(File.Exists(_destinationPath));
		}

		private int IncrementId(ref int id)
		{
			return id++;
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelWriterTests")]
		public void WriteFileUsingIDataReaderWithMaps()
		{
			var employees = Employee.GetTestEmployees(_employeesCount).ToList().GetDataReader();

			IList<IColumnInfoBase> employeesMap = new List<IColumnInfoBase>();
			//these expressions need to map to the column(field) names on the reader
			employeesMap.Add(new ColumnInfoBase("C", "L14All12Months", "(obj) => obj.Blah1"));
			employeesMap.Add(new ColumnInfoBase("D", "Line 15 Custom Header", "( obj ) => obj.Blah") { UpdateHeader = true }); //overwrite the current header
			employeesMap.Add(new ColumnInfoBase("E", "Line 16 Custom Header", "obj => obj.Blah2") { UpdateHeader = true }); //overwrite the current header

			using (ExcelWriter writer = new ExcelWriter(_sourcePath, 1, 2))
			{
				writer.WriteDataToSheet("Employees", employees, employeesMap);
				writer.WriteDataToSheet("Employer", Employer.GetEmployer().GetDataReader());

				writer.WriteTo(_destinationPath);
			}
			Assert.IsTrue(File.Exists(_destinationPath));
		}


		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelWriterTests")]
		public void WriteFileTestUsingHeaderNoMaps()
		{
			var employees = Employee.GetTestEmployees(_employeesCount).ToList();

			using (ExcelWriter writer = new ExcelWriter(_sourcePath, 1, 2))
			{
				writer.CreateSheetIfNotFound = true;
				writer.WriteDataToSheet("Employees", employees, null);
				writer.WriteDataToSheet("Employer", Employer.GetEmployer());
				writer.HeaderRow = 3;
				writer.DataRow = 5;
				writer.WriteDataToSheet("EmployeesGenerated", employees); //this sheet will auto gen with a different header and data row

				writer.WriteTo(_destinationPath);
			}
			Assert.IsTrue(File.Exists(_destinationPath));
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelWriterTests")]
		public void WriteFileTestUsingHeaderWithMaps()
		{
			var employees = Employee.GetTestEmployees(_employeesCount).ToList();
			ColumnInfoList<Employee> employeesMap = new ColumnInfoList<Employee>();
			//IList<IColumnInfo<Employee>> employeesMap = new List<IColumnInfo<Employee>>();
			employeesMap.Add("A", "", (obj) => String.Empty, updateHeader: true); //example on how to remove a columns data, does not actually remove the column
			employeesMap.Add("L14All12Months", (obj) => obj.Blah1);
			employeesMap.Add("L15All12Months", (obj) => obj.Blah);
			employeesMap.Add("L16All12Months", (obj) => obj.Blah2);
			employeesMap.Add("AT", "Name", (obj) => String.Format("{0} {1}", obj.L1Name, obj.L2SSN)); //add a new column outside of the headers at column AT
			employeesMap.Add("AU", "Name 2", (obj) => ComplicatedStringMethodExample(obj)); //next two are just show different ways to call the lambda
			employeesMap.Add("AV", "Name 3", (obj) => ConcatName(obj));
			employeesMap.Add(new ColumnInfo<Employee>("AW", "Hire Date", (obj) => obj.HireDate) { CellType = ExcelCellType.Date, FormatCode = "mm-dd-yyyy" });

			//new way
			using (ExcelWriter writer = new ExcelWriter(_sourcePath, 1, 2))
			{
				writer.WriteDataToSheet("Employees", employees, employeesMap);
				writer.WriteDataToSheet("Employer", Employer.GetEmployer());

				writer.WriteTo(_destinationPath);
			}
			Assert.IsTrue(File.Exists(_destinationPath));
		}

		private string ConcatName(Employee obj)
		{
			string name = obj.L1Name;
			string ssn = obj.L2SSN;
			return String.Format("{0} {1}", name, ssn);
		}

		private string ComplicatedStringMethodExample(Employee obj)
		{
			string name = obj.L1Name;
			string ssn = obj.L2SSN;
			return String.Format("{0} {1}", name, ssn);
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelWriterTests")]
		public void WriteFileTestNoHeaderCustomMappings()
		{
			var employees = Employee.GetTestEmployees(_employeesCount).ToList();
			ColumnInfoList<Employee> employeesMap = new ColumnInfoList<Employee>();

			//add our own weird custom mappings
			employeesMap.Add("A", "Name", (emp) => emp.L1Name);
			employeesMap.Add("C", "SSN", (emp) => emp.L2SSN.Replace("-", ""), cellType: ExcelCellType.Number, formatCode: "000-00-0000");
			employeesMap.Add("E", "Blah Col Set 1", (emp) => emp.Blah);
			employeesMap.Add("G", "Blah1 Col Set 1", (emp) => emp.Blah1);
			employeesMap.Add("I", "Blah2 Col Set 1", (emp) => emp.Blah2);
			employeesMap.Add("K", "L15 Jan Set 1", (emp) => emp.L15Jan, formatCode: "#,##0.00");
			employeesMap.Add("M", "L15 Jun Set 1", (emp) => Convert.ToDouble(emp.L15June) / 100d, cellType: ExcelCellType.Percent);
			employeesMap.Add("O", "Hire Date", (emp) => emp.HireDate, formatCode: "mm-dd-yyyy");
			employeesMap.Add("Q", "Hourly Rate", (emp) => GetRate(emp.HourlyRate), cellType: ExcelCellType.Currency, formatCode: "[Blue]$#,##0.00; [Red]-$#,##0.00;");

			employeesMap.Add("AA", "Col 1", (emp) => emp.L1Name);
			employeesMap.Add("AC", "Col 2", (emp) => emp.L2SSN);
			employeesMap.Add("AE", "Col 3", (emp) => emp.Blah);
			employeesMap.Add("AG", "Col 4", (emp) => emp.Blah1);
			employeesMap.Add("AI", "Col 5", (emp) => emp.Blah2);
			employeesMap.Add("AK", "Col 6", (emp) => PadLeft("000000", emp.L15Jan.ToString()), cellType: ExcelCellType.Quoted);
			employeesMap.Add("AM", "Col 7", (emp) => Convert.ToDouble(emp.L15June) / 100d, cellType: ExcelCellType.Percent, formatCode: "0.00%");
			employeesMap.Add("AO", "Col 8", (emp) => emp.HireDate);
			employeesMap.Add("AQ", "Col 9", (emp) => GetRate(emp.HourlyRate), cellType: ExcelCellType.Currency);


			ColumnInfoList<Employer> employerMap = new ColumnInfoList<Employer>();
			employerMap.Add(null, "L13Zip", (emp) => emp.L13Zip, cellType: ExcelCellType.Number);

			//setting a number less than 1 for the header row will cause the writer to generate the headers from the properties, so to completely jack with the format, lets override them all
			using (ExcelWriter writer = new ExcelWriter(1, 2))
			{
				writer.CreateSheetIfNotFound = true;
				writer.GenerateHeadersFromType = false; //we are going to lay out the sheet manually using the map 
				writer.WriteDataToSheet("Employees 2", employees, employeesMap);

				writer.GenerateHeadersFromType = true; //turn this back on for this sheet
				writer.WriteDataToSheet("Employer", Employer.GetEmployer(), employerMap);

				writer.WriteTo(_destinationPath);
			}

			Assert.IsTrue(File.Exists(_destinationPath));
		}

        private string PadLeft(string padding, string value)
        {
            var tmp = (padding + value);
            return tmp.Substring(Math.Min(tmp.Length - padding.Length, padding.Length));
        }

		private double GetRate(double rate)
		{
			if (rate <= 15.5d) { rate = -rate; }
			return rate;
		}


		// How to create custom validation.
		public static void ValidateSpreadSheet(string filepath)
		{
			try
			{

				OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2013);

				int count = 0;

				using (var spreadSheet = SpreadsheetDocument.Open(filepath, true))
				{
					foreach (ValidationErrorInfo error in validator.Validate(spreadSheet))
					{
						count++;
						Debug.WriteLine("Error " + count);
						Debug.WriteLine("Description: " + error.Description);
						Debug.WriteLine("Path: " + error.Path.XPath);
						Debug.WriteLine("Part: " + error.Part.Uri);
						Debug.WriteLine("——————————————-");

					}
				}
				Assert.AreEqual(0, count, string.Format("Found {0} validation errors", count));
				//Console.ReadKey();

			}

			catch (Exception ex)
			{

				Console.WriteLine(ex.Message);

			}
		}

	}
}
