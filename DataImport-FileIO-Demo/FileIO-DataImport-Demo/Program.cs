using FileIODemo.Data;
using FileIODemo.DataImport;
using FileIODemo.FileIO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace FileIODemo
{
    class Program
	{
		private static readonly object _lockObj = new object();

		static void Main(string[] args)
		{
			Common.Init();

			ConsoleKeyInfo response = new ConsoleKeyInfo();
			do
			{
				GC.Collect();
				GC.WaitForPendingFinalizers();
				var process = Process.GetCurrentProcess();
				long privateMem = process.PrivateMemorySize64 / 1024 / 1024;
				long workingMem = process.WorkingSet64 / 1024 / 1024;
				Console.WriteLine("Current Memory usage: \r\n\t{0} MB private memory \r\n\t{1} MB working set", privateMem, workingMem);

				Console.WriteLine();
				Console.WriteLine("[1] = Write Examples");
				Console.WriteLine("[2] = Read Examples");
				Console.WriteLine("[3] = Excel Write - Read Same File");
				Console.WriteLine("[4] = File Import Examples");
				Console.WriteLine("[5] = Looped Csv File Import");
				Console.WriteLine("[6] = Looped Excel File Import");
				Console.WriteLine("[7] = Enumerables Import");
				Console.WriteLine("[8] = Parallel Import");
				Console.WriteLine("[e] = Exit");

				Console.Write("\r\nEnter choice: ");
				response = Console.ReadKey();
				Console.WriteLine("\r\n");

				switch (response.Key)
				{
					case ConsoleKey.D1:
						/*** WRITE EXAMPLES ***/
						RunWriteExamples();
						Console.Clear();
						break;
					case ConsoleKey.D2:
						/*** READ EXAMPLES ***/
						RunReadExamples();
						Console.Clear();
						break;
					case ConsoleKey.D3:
						/*** READ-WRITE EXAMPLES ***/
						RunExcelWrite_ReadExamples();
						Console.Clear();
						break;
					case ConsoleKey.D4:
						/*** FILE IMPORT EXAMPLES ***/
						RunFileImportExamples();
						Console.Clear();
						break;
					case ConsoleKey.D5:
						/*** LOOPED CSV FILE IMPORT EXAMPLE ***/
						RunLoopedCsvFileImportExamples();
						Console.Clear();
						break;
					case ConsoleKey.D6:
						/*** LOOPED EXCEL FILE IMPORT EXAMPLE ***/
						RunLoopedExcelFileImportExamples();
						Console.Clear();
						break;
					case ConsoleKey.D7:
						/*** ENUMERABLE IMPORT EXAMPLE ***/
						RunEnumerablesImportExamples();
						Console.Clear();
						break;
					case ConsoleKey.D8:
						/*** PARALLEL ENUMERABLE IMPORT EXAMPLE ***/
						RunParallelImportExamples();
						Console.Clear();
						break;
					case ConsoleKey.E:
					case ConsoleKey.Escape:
						break;
					default:
						Console.WriteLine("Invalid choice.");
						break;
				}
			} while (!(response.Key == ConsoleKey.E || response.Key == ConsoleKey.Escape));

			//Console.WriteLine("\r\nDone.");
			//Console.ReadKey(true);
		}


		static void RunWriteExamples()
		{
			Console.WriteLine("RunWriteExamples()");

			var data = Common.GetCompanies();

			ConsoleKeyInfo response = new ConsoleKeyInfo();

			bool openFiles = false;
			Console.Write("\r\nDo you wish to open each file as it is created? [y|n] ");
			response = Console.ReadKey();
			openFiles = response.Key == ConsoleKey.Y;

			Console.WriteLine("\r\n");
			do
			{
				Console.WriteLine();
				Console.WriteLine("[1] = CsvIO.WriteCsvFile1");
				Console.WriteLine("[2] = CsvIO.WriteCsvFile2");
				Console.WriteLine("[3] = CsvIO.WriteCsvFile3");
				Console.WriteLine("[4] = ExcelIO.WriteExcelFile1");
				Console.WriteLine("[5] = ExcelIO.WriteExcelFile2");
				Console.WriteLine("[6] = ExcelIO.WriteExcelFile3");
				Console.WriteLine("[7] = ExcelIO.WriteExcelFile4");
				Console.WriteLine("[8] = ExcelIO.WriteExcelFile5");
				Console.WriteLine("[e] = Exit");

				Console.Write("\r\nEnter choice: ");
				response = Console.ReadKey();
				Console.WriteLine("\r\n");

				Stopwatch sw = null;
				string path = null;

				switch (response.Key)
				{
					case ConsoleKey.D1:
						sw = Stopwatch.StartNew();
						path = CsvIO.WriteCsvFile1(data);
						break;
					case ConsoleKey.D2:
						sw = Stopwatch.StartNew();
						path = CsvIO.WriteCsvFile2(data);
						break;
					case ConsoleKey.D3:
						sw = Stopwatch.StartNew();
						path = CsvIO.WriteCsvFile3();
						break;
					case ConsoleKey.D4:
						sw = Stopwatch.StartNew();
						path = ExcelIO.WriteExcelFile1(data);
						break;
					case ConsoleKey.D5:
						sw = Stopwatch.StartNew();
						path = ExcelIO.WriteExcelFile2(data);
						break;
					case ConsoleKey.D6:
						sw = Stopwatch.StartNew();
						path = ExcelIO.WriteExcelFile3(data);
						break;
					case ConsoleKey.D7:
						sw = Stopwatch.StartNew();
						path = ExcelIO.WriteExcelFile4(data);
						break;
					case ConsoleKey.D8:
						sw = Stopwatch.StartNew();
						path = ExcelIO.WriteExcelFile5();
						break;
					case ConsoleKey.E:
					case ConsoleKey.Escape:
						break;
					default:
						Console.WriteLine("Invalid choice.");
						break;
				}

				if (sw != null && sw.IsRunning)
				{
					sw.Stop();
					Common.WriteTime("File write operation total seconds: ", sw.Elapsed);
					Console.WriteLine("File written to: {0}", path);
					if (openFiles)
					{
						Process.Start(path);
					}
				}
			} while (!(response.Key == ConsoleKey.E || response.Key == ConsoleKey.Escape));

			data = null;
		}



		static void RunReadExamples()
		{
			Console.WriteLine("RunReadExamples()");


			IEnumerable<Company> companies = Enumerable.Empty<Company>();

			ConsoleKeyInfo response = new ConsoleKeyInfo();

			do
			{
				Console.WriteLine();
				Console.WriteLine("[1] = CsvIO.ReadCsvFile1");
				Console.WriteLine("[2] = ExcelIO.ReadExcelFile1");
				Console.WriteLine("[e] = Exit");

				Console.Write("\r\nEnter choice: ");
				response = Console.ReadKey();
				Console.WriteLine("\r\n");

				Stopwatch sw = null;

				switch (response.Key)
				{
					case ConsoleKey.D1:
						sw = Stopwatch.StartNew();
						companies = CsvIO.ReadCsvFile1();
						break;
					case ConsoleKey.D2:
						sw = Stopwatch.StartNew();
						companies = ExcelIO.ReadExcelFile1();
						break;
					case ConsoleKey.E:
					case ConsoleKey.Escape:
						break;
					default:
						Console.WriteLine("Invalid choice.");
						break;
				}

				if (sw != null && sw.IsRunning)
				{
					sw.Stop();
					Common.WriteTime(String.Format("\r\nRead file returning {0} companies total seconds: ", companies.Count()), sw.Elapsed);
				}
			} while (!(response.Key == ConsoleKey.E || response.Key == ConsoleKey.Escape));

		}



		static void RunExcelWrite_ReadExamples()
		{
			Console.WriteLine("RunExcelWrite_ReadExamples()");

			var data = Common.GetCompanies();

			Console.WriteLine("\r\nFile write started");
			Stopwatch sw = Stopwatch.StartNew();

			//write out a file
			string path = ExcelIO.WriteExcelFile1(data);

			sw.Stop();
			Common.WriteTime("File write operation total seconds: ", sw.Elapsed);

			sw = Stopwatch.StartNew();
			Console.WriteLine("\r\nFile read started");

			//read that same file written out back in
			var companies = ExcelIO.ReadExcelFile1(path, 1, 2);

			sw.Stop();
			Common.WriteTime(String.Format("\r\nRead file returning {0} companies total seconds: ", companies.Count), sw.Elapsed);

			data = null;
			Console.WriteLine("File written to: {0}", path);
			Console.WriteLine("\r\nPress any key to continue.");
			Console.ReadKey(true);
		}



		static void RunFileImportExamples()
		{
			Console.WriteLine("RunFileImportExamples()");

			ConsoleKeyInfo response = new ConsoleKeyInfo();

			do
			{
				Console.WriteLine();
				Console.WriteLine("[1] = CsvImport1");
				Console.WriteLine("[2] = ExcelImport1");
				Console.WriteLine("[e] = Exit");

				Console.Write("\r\nEnter choice: ");
				response = Console.ReadKey();
				Console.WriteLine("\r\n");

				Stopwatch sw = null;

				switch (response.Key)
				{
					case ConsoleKey.D1:
						sw = Stopwatch.StartNew();
						CsvDataImport.CsvImport1(Common.CsvDataPath);
						break;
					case ConsoleKey.D2:
						sw = Stopwatch.StartNew();
						ExcelDataImport.ExcelImport1(Common.ExcelDataPath);
						break;
					case ConsoleKey.E:
					case ConsoleKey.Escape:
						break;
					default:
						Console.WriteLine("Invalid choice.");
						break;
				}

				if (sw != null && sw.IsRunning)
				{
					sw.Stop();
					Common.WriteTime("Import total seconds: ", sw.Elapsed);
				}
			} while (!(response.Key == ConsoleKey.E || response.Key == ConsoleKey.Escape));


		}



		static void RunLoopedCsvFileImportExamples()
		{
			Console.WriteLine("RunLoopedCsvFileImportExamples()");

			Console.WriteLine("Import started");

			for (int i = 0; i < 5; i++)
			{
				Stopwatch sw = Stopwatch.StartNew();

				CsvDataImport.CsvImport1(Common.CsvDataPath);

				sw.Stop();
				Common.WriteTime(String.Format("Import {0} total seconds: ", i), sw.Elapsed);
			}
			Console.WriteLine("\r\nPress any key to continue.");
			Console.ReadKey(true);
		}



		static void RunLoopedExcelFileImportExamples()
		{
			Console.WriteLine("RunLoopedExcelFileImportExamples()");

			Console.WriteLine("Import started");

			for (int i = 0; i < 5; i++)
			{
				Stopwatch sw = Stopwatch.StartNew();

				ExcelDataImport.ExcelImport1(Common.ExcelDataPath);

				sw.Stop();
				Common.WriteTime(String.Format("Import {0} total seconds: ", i), sw.Elapsed);
			}
			Console.WriteLine("\r\nPress any key to continue.");
			Console.ReadKey(true);
		}



		static void RunEnumerablesImportExamples()
		{
			Console.WriteLine("RunEnumerablesImportExamples()");

			Console.WriteLine("File read started");
			Stopwatch sw = Stopwatch.StartNew();

			var companies = ExcelIO.ReadExcelFile1().ToList();
			//var companies = ExcelIO.ReadExcelFile2(Common.ExcelDataPath, 2, 3).ToList();

			sw.Stop();
			Common.WriteTime(String.Format("\r\nRead file returning {0} companies total seconds: ", companies.Count), sw.Elapsed);

			//Using a file here to get the enumerable, but in a normal application this would come from any collection
			Console.WriteLine("\r\nImport started");
			sw = Stopwatch.StartNew();

			EnumerablesDataImport.EnumerableImport1(companies);
			//EnumerablesDataImport.EnumerableImport2(companies);

			sw.Stop();
			Common.WriteTime("Import total seconds: ", sw.Elapsed);
			Console.WriteLine("\r\nPress any key to continue.");
			Console.ReadKey(true);
		}



		static void RunParallelImportExamples()
		{
			/* This example demonstrates that multiple inserts can hit the same destination table at the same time. 
			 * While the insert process slows down, they do run. The slowdown is to be expected. 
			 * 
			 * Normally I would NOT recommend doing parallelism for data imports. This however demonstrates that multiple 
			 * imports could be instantiated from different users of a web site.
			 */
			Console.WriteLine("RunParallelImportExamples()");

			var companies = Common.GetCompanies().ToList();

			Console.WriteLine("\r\nImport started");
			Stopwatch sw = Stopwatch.StartNew();

			//apply a max threads count
			var options = new ParallelOptions() { MaxDegreeOfParallelism = 8 };

			int loopCount = 50;
			double totalSeconds = 0;
			double minSeconds = 99999;
			double maxSeconds = 0;
			
			Parallel.For(0, loopCount, options, (x) =>
			{
				Console.WriteLine("Import {0} started", x);
				Stopwatch sw2 = Stopwatch.StartNew();

				//CsvDataImport.CsvImport1(Common.CsvDataPath);
				//cant parallel the same excel file. The first time the file is opened will cause the file to be locked
				//ExcelDataImport.ExcelImport1(Common.ExcelDataPath);
				EnumerablesDataImport.EnumerableImport2(companies.ToArray());

				sw2.Stop();
				Common.WriteTime(String.Format("Import {0} finished. total seconds: ", x), sw2.Elapsed);
				
				//lock this code area so multiple threads can not run it at the same time
				lock (_lockObj)
				{
					minSeconds = Math.Min(minSeconds, sw2.Elapsed.TotalSeconds);
					maxSeconds = Math.Max(maxSeconds, sw2.Elapsed.TotalSeconds);
					totalSeconds += sw2.Elapsed.TotalSeconds;
				}
			});

			sw.Stop();
			Console.WriteLine("\r\nSummary:");
			Common.WriteMessage("Minimum (seconds): ", minSeconds);
			Common.WriteMessage("Maximum (seconds): ", maxSeconds);
			Common.WriteMessage("Average (seconds): ", totalSeconds / Convert.ToDouble(loopCount));
			Common.WriteMessage("Total (seconds): ", sw.Elapsed.TotalSeconds);
			Common.WriteMessage("Row Count: ", String.Format("{0:n0}", loopCount * companies.Count()));
			Console.WriteLine("\r\nPress any key to continue.");
			Console.ReadKey(true);
		}

	}
}
