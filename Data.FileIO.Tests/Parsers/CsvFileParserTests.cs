using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Data.FileIO.Common;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Core;
using Data.FileIO.Interfaces;
using Data.FileIO.Parsers;
using Data.FileIO.Writers;
using Data.FileIOTests.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace Data.FileIO.Parsers.Tests
{
	[TestClass(), DeploymentItem("Files\\1095_Import_Employees.csv", "Files")]
	public class CsvFileParserTests
	{
		public TestContext TestContext { get; set; }

		[TestMethod(), TestCategory("Unit"), TestCategory("FileIO_CsvFileParserTests")]
		public void ParseCsvFileTest()
		{

			string filePath = Path.Combine(this.TestContext.TestDeploymentDir, "Files\\1095_Import_Employees.csv");

			ICsvFileParser parser = new CsvFileParser(filePath);
			IDictionary<int, IList<string>> fileErrors = new Dictionary<int, IList<string>>();
			IObjectValidator validator = new ObjectValidator();

			int rowIndex = parser.RowStart;

			foreach (dynamic row in parser.ParseFile())
			{
				List<string> errors = new List<string>();

				Employee rowObj = FileIOUtilities.MapObject<Employee>(row, rowIndex, validator, null, ref errors);
				//rowObj.MapValues(rowIndex, row, validator, ref errors);
				validator.TryValidate(rowObj, ref errors);

				if (errors.Count > 0)
				{
					fileErrors.Add(rowIndex, errors);
				}
				rowIndex++;
			}

			Assert.IsTrue(fileErrors.Count >= 2);
		}
	}
}
