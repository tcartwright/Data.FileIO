using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Data.FileIO.Common;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Interfaces;
using Data.FileIOTests.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace Data.FileIO.Parsers.Tests
{
	[TestClass(), DeploymentItem("Files\\1095_Import_WithData.xlsx", "Files")]
	public class ExcelFileParserTests
	{
		public TestContext TestContext { get; set; }

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
		public void ParseExcelFileTest()
		{

			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);
			IDictionary<int, IList<string>> fileErrors = new Dictionary<int, IList<string>>();
			IObjectValidator validator = new ObjectValidator();

			int rowIndex = parser.RowStart;

			foreach (dynamic row in parser.ParseFile())
			{
				List<string> errors = new List<string>();

				Employee rowObj = new Employee();
				rowObj.MapValues(rowIndex, row, validator, ref errors);
				validator.TryValidate(rowObj, ref errors);

				if (errors.Count > 0)
				{
					fileErrors.Add(rowIndex, errors);
				}
				rowIndex++;
			}

			Assert.IsTrue(fileErrors.Count >= 2);

			parser.SheetName = "Dependents";
			fileErrors = new Dictionary<int, IList<string>>();
			rowIndex = parser.RowStart;

			foreach (dynamic row in parser.ParseFile())
			{
				List<string> errors = new List<string>();

				Dependent rowObj = new Dependent();
				rowObj.MapValues(rowIndex, row, validator, ref errors);
				validator.TryValidate(rowObj, ref errors);

				if (errors.Count > 0)
				{
					fileErrors.Add(rowIndex, errors);
				}
				rowIndex++;
			}

			Assert.IsTrue(fileErrors.Count == 0);
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
		public void ParseFileHeaders_AsIsTest()
		{
			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);

			var headers = parser.GetHeaders();

			Assert.AreEqual(42, headers.Count);
			Assert.AreEqual("L1. Name", headers["A"]);
			Assert.AreEqual("L2. SSN", headers["B"]);
		}

        [TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
        public void ParseSheetNames_AsIsTest()
        {
            string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");


            var sheetNames = ExcelFileParser.GetSheetNames(filePath);

            Assert.AreEqual(4, sheetNames.Count);
        }

        [TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
		public void ParseFileHeaders_NoBadCharsTest()
		{
			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);

			var headers = parser.GetHeaders(true);

			Assert.AreEqual(42, headers.Count);
			Assert.AreEqual("L1Name", headers["A"]);
			Assert.AreEqual("L2SSN", headers["B"]);
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
		public void ParseFileHeaders_StopAtColumn()
		{
			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);

			parser.EndColumnKey = "H";

			var headers = parser.GetHeaders(true);

			Assert.AreEqual(8, headers.Count);
			Assert.AreEqual("L1Name", headers["A"]);
			Assert.AreEqual("L2SSN", headers["B"]);
		}

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_ExcelFileParserTests")]
		public void ParseFile_StopAtColumn()
		{
			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_WithData.xlsx");

			IExcelFileParser parser = new ExcelFileParser(filePath, "Employees", 1, 2);
			IDictionary<int, IList<string>> fileErrors = new Dictionary<int, IList<string>>();
			IObjectValidator validator = new ObjectValidator();

			parser.EndColumnKey = "H";

			var data = parser.ParseFile();
			dynamic firstRow = data.First();

			Assert.AreEqual("Lucy Davis", firstRow.L1Name);
			Assert.AreEqual("308-35-1715", firstRow.L2SSN);

		}

	}
}
