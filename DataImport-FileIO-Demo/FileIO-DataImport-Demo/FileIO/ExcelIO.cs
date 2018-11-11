using Data.FileIO;
using Data.FileIO.Common;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Core;
using Data.FileIO.Interfaces;
using Data.FileIO.Parsers;
using Data.FileIO.Writers;
using FileIODemo.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace FileIODemo.FileIO
{
    public static class ExcelIO
	{
		#region excel writers
		
		public static string WriteExcelFile1(IEnumerable<Company> companies)
		{
			Console.WriteLine("Running WriteExcelFile1");
			/*This simple example shows writing a Company class out to a file.  */

			var path = Common.GetFileName(Common.XLS_TYPE);
			using (var writer = new ExcelWriter(1, 2)) //header row, and data row. Could be different per sheet
			{
				//write the collection to a sheet
				writer.WriteDataToSheet("Companies", companies);
				writer.WriteTo(path); //WriteTo can write to a path, OR a stream so it can be used using Repsonse.Write
			}

			return path;
		}
		
		public static string WriteExcelFile2(IEnumerable<Company> companies)
		{
            //companies = companies.Take(5000);
			Console.WriteLine("Running WriteExcelFile2");
            /*This example shows writing a Company class a sheet using a simple map.  */

            var map = new ColumnInfoList<Company>
            {
                { "A", "Company Id", obj => obj.CompanyId },
                { "B", "Legal Name", obj => obj.LegalName },
                { "D", "DBA Name", obj => obj.DBAName },
                { "E", "Change Date", obj => obj.ChangeDate }
            };


            var path = Common.GetFileName(Common.XLS_TYPE);
			using (var writer = new ExcelWriter(1, 2))
			{
				writer.WriteDataToSheet("Companies", companies, map);
				//showing appending to a sheet
				writer.WriteDataToSheet("Companies", companies, map, true);

				writer.WriteTo(path);
			}

			return path;
		}
		
		public static string WriteExcelFile3(IEnumerable<Company> companies)
		{
			Console.WriteLine("Running WriteExcelFile3");
			/*This example shows writing a Company class to multiple sheets without using maps. The Companies sheet uses a template.  */

			var path = Common.GetFileName(Common.XLS_TYPE);
			using (var writer = new ExcelWriter(Common.TemplatePath, 2, 3))
			{
				writer.CreateSheetIfNotFound = true; //some of the sheets do not exist in the template, so its important we tell the writer we can create new sheets
				writer.WriteDataToSheet("Companies", companies);

				writer.HeaderRow = 3;  //optionally adjust the header, and footer row per sheet.
				writer.DataRow = 4;
				writer.WriteDataToSheet("Companies2", companies);

				writer.HeaderRow = 4; //optionally adjust the header, and footer row per sheet.
				writer.DataRow = 5;
				writer.WriteDataToSheet("Another Companies  Sheet", companies);
				writer.WriteTo(path);
			}

			return path;
		}
		
		public static string WriteExcelFile4(IEnumerable<Company> companies)
		{
			Console.WriteLine("Running WriteExcelFile4");
			/*This example shows writing a Company class to a file with a column mapping that writes the columns in order as they appear to the sheet */

			//, formatCode: "#,##0.00"
			//, cellType: ExcelCellType.Currency, formatCode: "[Blue]$#,##0.00; [Red]-$#,##0.00;"
			//, cellType: ExcelCellType.Percent, formatCode: "0.00%"
			var map = new ColumnInfoList<Company>();
			//the column type is inferred from the object type
			map.Add("Company Id", (c) => c.CompanyId, cellType: ExcelCellType.Number);
			map.Add("Legal Name", (c) => c.LegalName);
			//if updateheader is true we will update the header value. In this case WE have to pass the column name
			map.Add("E", "Doing Business As", (c) => c.DBAName, updateHeader: true);
			//specialiazed format codes can be used to format the data as desired.
			//	NOTE: IF you are writing to an existing template, that already has the column formatted then this formatcode will be ignored
			map.Add("Change Date", (c) => c.ChangeDate, cellType: ExcelCellType.Date, formatCode: "mm-dd-yyyy");
			map.Add("UserId", (c) => c.UserId);

			var path = Common.GetFileName(Common.XLS_TYPE);
			using (var writer = new ExcelWriter(Common.TemplatePath, 2, 3))
			{
				writer.CreateSheetIfNotFound = true;
				writer.WriteDataToSheet("Companies", companies, map);
				writer.WriteTo(path);
			}

			return path;
		}
		
		public static string WriteExcelFile5()
		{
			Console.WriteLine("Running WriteExcelFile5");
			/*This example shows writing a Company class to a file with a column mapping  that provides total control over the file layout and uses a iDataReader for the data source */

			Func<int, string> col = (i) => FileIOUtilities.ConvertExcelNumberToLetter(i);
			var y = 1;

            //, formatCode: "#,##0.00"
            //, cellType: ExcelCellType.Currency, formatCode: "[Blue]$#,##0.00; [Red]-$#,##0.00;"
            //, cellType: ExcelCellType.Percent, formatCode: "0.00%"
            var map = new ColumnInfoList<Company>
            {
                { col(y++), "Company Id", (c) => c.CompanyId },
                //insert a blank column at B for Foo, next column starts at C
                { col(y++), "Foo", (c) => null },  //, formatCode: "000-00-0000"
                { col(y++), "Legal Name", (c) => c.LegalName }
            };
            map.Add(col(y++), "Doing Business As", (c) => c.DBAName, updateHeader: true);
			map.Add(col(y++), "Change Date", (c) => c.ChangeDate, cellType: ExcelCellType.Date, formatCode: "mm-dd-yyyy");
			map.Add(col(y++), "UserId", (c) => c.UserId);

			var path = Common.GetFileName(Common.XLS_TYPE);
			using (var writer = new ExcelWriter(1, 3))
			{
				writer.CreateSheetIfNotFound = true;
				
				using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataModel"].ConnectionString))
				{
					conn.Open();

					using (var cmd = conn.CreateCommand())
					{
						cmd.CommandText = "Select top (10000) * from Company";
						cmd.CommandType = CommandType.Text;

						using (var reader = cmd.ExecuteReader())
						{
							writer.WriteDataToSheet("Companies", reader, map);
							writer.WriteTo(path);
						}
					}
				}

			}

			return path;
		}


		#endregion excel writers

		#region excel readers
		public static IList<Company> ReadExcelFile1(string path = null, int headerRow = 2, int dataRow = 3)
		{
			/* Some important notes about header names:
			 *	1) As the header name is being used for the "property" name of the dynamic object there are many characters that are not valid to be a valid c# property name. 
			 *	During the reading of the file, the invalid characters in the header name are replaced using this regex [^A-Za-z0-9_]* with an empty string. 
			 *	Then if the return of that replace operation is empty, then the column is named "Col(i)". Where (i) is the zero based column index.
			 *	Examples:
			 *	Column Name---------------Property Name
			 *	Company Id----------------row.CompanyId
			 *	Aims Company Id-----------row.AimsCompanyId
			 *	Some_Id_9-----------------row.Some_Id_9
			 *	(@@@)---------------------row.Col1 - where 1 is the zero based index of that column. 
			 *	
			 *  2) Case sensitivity of the property names does not matter either. As the sender could change the case indiscrimiately.
			 *  3) If a column is removed that you were expecting, the property will return empty and will not throw an exception.
			 */


			Console.WriteLine("Running ReadExcelFile1");
			/*This example shows reading a Company class from a file.  */

			if (String.IsNullOrWhiteSpace(path)) { path = Common.ExcelDataPath; }
			//the parser is designed to read one sheet from a file at a time. Other sheets would require a new parser, or just use the same parser and change the sheet name.
			IExcelFileParser parser = new ExcelFileParser(path, "Companies", headerRow, dataRow);
			//this is where we will store any parser errors as we parse the file
			var fileErrors = new Dictionary<int, IList<string>>();
			//this validator provides validation when parsing the columns, and data attributes like [Required] and it will also invoke the IValidatableObject  interface 
			IObjectValidator validator = new ObjectValidator();

			//as the rowindex may not start at row 1, get it from the parser
			int rowIndex = parser.RowStart;
			var companies = new List<Company>();

			foreach (dynamic row in parser.ParseFile())
			{
				List<string> errors = new List<string>();

				//create a reference to a custom mapper. this provides complete control over the mapping of the file row to the object, and the interface is skipped
				FileValuesMap<Company> mapper = Common.Company_FileMapper;

				//this utility performs mapping of the row to the object and invokes column mapping validation
				var rowObj = FileIOUtilities.MapObject<Company>(row, rowIndex, validator, mapper, ref errors);
				//calling the TryValidate method invoke the data annotation validation, and the IValidatableObject  interface 
				validator.TryValidate(rowObj, ref errors);

				companies.Add(rowObj);

				if (errors.Count > 0)
				{
					//we got errors for this row, so add them to the fileErrors dictionary
					fileErrors.Add(rowIndex, errors);
				}
				rowIndex++;
			}

			//write all the file errors out to the console
			foreach (var errs in fileErrors)
			{
				foreach (var err in errs.Value)
				{
					Console.WriteLine("Line:{0}, Error: {1}", errs.Key, err);
				}
			}

			return companies;
		}

		#endregion excel readers
	}
}
