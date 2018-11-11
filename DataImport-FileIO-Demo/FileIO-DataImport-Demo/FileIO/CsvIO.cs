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
    static class CsvIO
	{
		#region csv writers
		public static string WriteCsvFile1(IEnumerable<Company> companies)
		{
			Console.WriteLine("Running WriteCsvFile1");

			string path = Common.GetFileName(Common.CSV_TYPE);
			//writes the csv file using quotes values
			CsvWriter writer = new CsvWriter();
			writer.WriteData(path, companies, '"');
			return path;
		}
		
		public static string WriteCsvFile2(IEnumerable<Company> companies)
		{
			Console.WriteLine("Running WriteCsvFile2");

			var i = 1;
			var path = Common.GetFileName(Common.CSV_TYPE);
            //writes the csv file, mapping all of the columns / data
            var map = new CsvColumnInfoList<Company>
            {
                { "Id", obj => i++ },
                { "Company Id", obj => obj.CompanyId },
                { "Legal Name", obj => obj.LegalName, '"' }, //set a quoted identifier on this field only
                { "DBA Name", obj => obj.DBAName, '"' }, //set a quoted identifier on this field only
                { "Change Date", obj => obj.ChangeDate }
            };

            CsvWriter writer = new CsvWriter();
			writer.WriteData(path, companies, map);
			return path;
		}

		public static string WriteCsvFile3()
		{
			Console.WriteLine("Running WriteCsvFile3");

            var i = 1;
            var path = Common.GetFileName(Common.CSV_TYPE);
			var writer = new CsvWriter();

            //writes the csv file, mapping all of the columns / data
            var map = new CsvColumnInfoList<Company>
            {
                { "Id", obj => i++ },
                { "Company Id", obj => obj.CompanyId },
                { "Legal Name", obj => obj.LegalName, '"' }, //set a quoted identifier on this field only
                { "DBA Name", obj => obj.DBAName, '"' }, //set a quoted identifier on this field only
                { "Change Date", obj => obj.ChangeDate }
            };

            using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataModel"].ConnectionString))
			{
				conn.Open();

				using (var cmd = conn.CreateCommand())
				{
					cmd.CommandText = "Select top (10000) * from Company";
					cmd.CommandType = CommandType.Text;

					using (var reader = cmd.ExecuteReader())
					{
						writer.WriteData(path, reader, map);
					}
				}
			}
			return path;
		}

		#endregion csv writers


		#region csv readers
		public static IList<Company> ReadCsvFile1()
		{
			Console.WriteLine("Running ReadCsvFile1");
			var parser = new CsvFileParser(Common.CsvDataPath);
			var fileErrors = new Dictionary<int, IList<string>>();
			var validator = new ObjectValidator();
			
			//parser.Delimiters // can adjust the delimiters
			//parser.FixedWidths // can parse fixed widths

			var rowIndex = parser.RowStart;
			var companies = new List<Company>();

			foreach (dynamic row in parser.ParseFile())
			{
				var errors = new List<string>();

				var rowObj = FileIOUtilities.MapObject<Company>(row, rowIndex, validator, null, ref errors);
				validator.TryValidate(rowObj, ref errors);
				companies.Add(rowObj);

				if (errors.Count > 0)
				{
					fileErrors.Add(rowIndex, errors);
				}
				rowIndex++;
			}

			foreach (var errs in fileErrors)
			{
				foreach (var err in errs.Value)
				{
					Console.WriteLine("Line:{0}, Error: {1}", errs.Key, err);
				}
			}

			return companies;
		}
		#endregion csv readers

	}
}
