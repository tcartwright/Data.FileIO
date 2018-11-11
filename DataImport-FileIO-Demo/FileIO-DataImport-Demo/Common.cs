using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using FileIODemo.Data;
using Data.DataImport;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Microsoft.SqlServer.Server;

namespace FileIODemo
{
	static class Common
	{
		private static readonly object _lockObj = new object();

		public const string CSV_TYPE = ".csv";
		public const string XLS_TYPE = ".xlsx";

		public static readonly DateTime MinSqlServerDate = new DateTime(1753, 1, 1);
		public static readonly DateTime MaxSqlServerDate = new DateTime(9999, 12, 31);

		public static string TemplatePath { get; private set; }
		public static string CsvDataPath { get; private set; }
		public static string ExcelDataPath { get; private set; }

		public static void Init()
		{
			string exedir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			TemplatePath = Path.Combine(exedir, "Misc", "Templates", "Simple Template.xlsx");
			CsvDataPath = Path.Combine(exedir, "Misc", "DataFiles", "CompaniesData.csv");
			ExcelDataPath = Path.Combine(exedir, "Misc", "DataFiles", "CompaniesData.xlsx");

			//set the data directory to use for when connecting to the local db
			AppDomain.CurrentDomain.SetData("DataDirectory", exedir);

		}

		public static IQueryable<Company> GetCompanies(int count = 10000)
		{
			Stopwatch sw = Stopwatch.StartNew();
			Console.WriteLine("Data retrieval started");
			var dm = new DataModel();
			var data = dm.Company.Take(count);
			sw.Stop();
			WriteTime("Data retrieval complete, total seconds: ", sw.Elapsed);

			return data;
		}

		public static void WriteTime(string message, TimeSpan span)
		{
			//try to keep the threads test from bumping into each other when writing out their message
			lock (_lockObj) 
			{
				Console.Write(message);
				Console.ForegroundColor = ConsoleColor.Yellow;
				Console.WriteLine(span.TotalSeconds);
				Console.ResetColor();
			}
		}

		public static void WriteMessage(string message, object value)
		{
			//try to keep the threads test from bumping into each other when writing out their message
			lock (_lockObj)
			{
				Console.Write(message);
				Console.ForegroundColor = ConsoleColor.Yellow;
				Console.WriteLine(value);
				Console.ResetColor();
			}
		}

		public static string GetFileName(string ext)
		{
			string dir = Path.Combine(Path.GetTempPath(), "FileIOExamples");
			Directory.CreateDirectory(dir);
			string filename = Path.GetFileName(Path.GetTempFileName() + ext);
			return Path.Combine(dir, filename);
		}

		public static void MapValues(Company mapObject, SqlDataRecord record, int rowIndex, IEnumerable<string> errors)
		{
			record.SetInt32("CompanyId", mapObject.CompanyId);
			SetDateTime(record, "StartDate", mapObject.StartDate);
			SetDateTime(record, "EndDate", mapObject.EndDate);
			record.SetString("LegalName", mapObject.LegalName);
			record.SetString("DBAName", mapObject.DBAName);
			SetDateTime(record, "ChangeDate", mapObject.ChangeDate);
			record.SetString("UserId", mapObject.UserId);

			if (errors.Count() > 0)
			{
				record.SetString("ImportErrors", FileIOHelpers.ErrorsToXml(errors, rowIndex));
			}
		}

		#region Custom mapper
		/// <summary>
		/// Maps the values from each row in the file to a class in the enumerable.
		/// </summary>
		/// <param name="mapObject">The map object.</param>
		/// <param name="rowIndex">Index of the row.</param>
		/// <param name="row">The row coming in from the file. A Dynamic object, with a property per column</param>
		/// <param name="validator">The validator.</param>
		/// <param name="errors">The errors.</param>
		public static void Company_FileMapper(ref Company mapObject, int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors)
		{
			mapObject.CompanyId = validator.GetRowValue<int>(row, "CompanyId", ref errors);
			mapObject.StartDate = validator.GetRowValue<DateTime>(row, "StartDate", ref errors);
			mapObject.EndDate = validator.GetRowValue<DateTime>(row, "EndDate", ref errors, isNullable: true);
			mapObject.LegalName = row.LegalName;
			mapObject.DBAName = row.DBAName;
			mapObject.ChangeDate = validator.GetRowValue<DateTime>(row, "ChangeDate", ref errors);
			mapObject.UserId = row.UserId;
		}

		#endregion

		public static void SetDateTime (SqlDataRecord record, string key, DateTime? date)
		{
			//choices:
			// we could A) Ignore the invalid date, and pass it in like below. This lets the errors get recorded into the db or log...
			//			B) Throw an exception
			//			C) Pass it in, sql throws an exception


			//check the date to make sure it is valid for sql server, as parsing an invalid date will leave datetimes as datetime.minvalue 
			if (date.HasValue)
			{
				if (date.Value > Common.MinSqlServerDate && date.Value < Common.MaxSqlServerDate)
				{
					record.SetDateTime(key, date);
				}
				else
				{
					Debug.WriteLine("Invalid date");
				}
			}
		}
	}
}
