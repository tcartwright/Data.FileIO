using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport;
using Data.DataImport.Interfaces;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Interfaces;

namespace Data.FileIOTests.TestSupport
{
	public class Employee : IFileRowMapper, ISqlRecordMapper
	{
		private static string[] _codes1 = { "1A", "1B", "1C", "1D", "1E", "1F", "1G", "1H", "1I" };
		private static string[] _codes2 = { "2A", "2B", "2C", "2D", "2E", "2F", "2G", "2H", "2I" };

		[Required(AllowEmptyStrings = false)]
		public string L2SSN { get; set; }
		[Required(AllowEmptyStrings = false)]
		public string L1Name { get; set; }
		public string Blah { get; set; }
		public string Blah1 { get; set; }
		public string Blah2 { get; set; }
		public int L15Jan { get; set; }
		public int L15June { get; set; }

		public DateTime HireDate { get; set; }

		public double HourlyRate { get; set; }

		#region faker test data
		public static IEnumerable<Employee> GetTestEmployees(int employeesCount)
		{
			DateTime min = DateTime.Now.AddYears(-10);
			Random rand = new Random();
			for (int i = 0; i < employeesCount; i++)
			{
				yield return new Employee()
				{
					L2SSN = GetFakeSSN(),
					L1Name = Faker.NameFaker.Name(),
					HireDate = Faker.DateTimeFaker.DateTime(min, DateTime.Now),
					HourlyRate = Convert.ToDouble(Faker.NumberFaker.Number(10, 20)) + rand.NextDouble(),
					Blah = "L15All12Months " + Faker.StringFaker.AlphaNumeric(5),
					Blah1 = Faker.ArrayFaker.SelectFrom(_codes1),
					Blah2 = Faker.ArrayFaker.SelectFrom(_codes2),
					L15Jan = Faker.NumberFaker.Number(1, 99),
					L15June = Faker.NumberFaker.Number(1, 99)
				};
			}

		}

		private static string GetFakeSSN()
		{
			Func<int> rand1 = () => Faker.NumberFaker.Number(0, 9);
			Func<int> rand2 = () => Faker.NumberFaker.Number(1, 9);

			return String.Format("{0}{1}{2}-{3}{4}-{5}{6}{7}{8}"
				, rand1(), rand2(), rand2()
				, rand1(), rand2()
				, rand1(), rand2(), rand2(), rand2());
		}
		#endregion faker test data

		#region IFileRowMapper Members

		public void MapValues(int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors)
		{
			this.L2SSN = row.L2SSN;
			this.L1Name = row.L1Name;
			this.L15Jan = validator.GetRowValue<int>(row, "L15Jan", ref errors);
			this.L15June = validator.GetRowValue<int>(row, "L15June", ref errors);
			this.Blah = row.L15All12Months;
			this.Blah1 = row.L14All12Months;
			this.Blah2 = row.L16All12Months;
		}

		#endregion

		#region ISqlRecordMapper Members

		public void MapSqlRecord(Microsoft.SqlServer.Server.SqlDataRecord record, int rowIndex, IEnumerable<string> errors)
		{
			record.SetString("SSN", this.L2SSN);
			record.SetString("Name", this.L1Name);
			string errs = errors.Count() > 0 ? FileIOHelpers.ErrorsToXml(errors, rowIndex) : null;
			record.SetString("Errors", errs);
		}


		#endregion
	}


}
