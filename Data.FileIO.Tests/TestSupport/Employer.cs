using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data.FileIOTests.TestSupport
{
	public class Employer
	{
		private static string[] _states = { "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY" };

		public string L7Name { get; set; }
		public int L8EIN { get; set; }
		public string L9StreetAddress { get; set; }
		public string L10Phone { get; set; }
		public string L11City { get; set; }
		public string L12State { get; set; }
		public string L13Zip { get; set; }
		public string PlanStartMonth { get; set; }


		public static IEnumerable<Employer> GetEmployer()
		{
			string month = "0" + Convert.ToString(Faker.NumberFaker.Number(1, 12));
			month = month.Substring(month.Length - 2);

			yield return new Employer()
			{
				L7Name = Faker.CompanyFaker.Name(),
				L8EIN = Faker.NumberFaker.Number(1000000, 9999999),
				L10Phone = Faker.PhoneFaker.Phone(),
				L11City = Faker.LocationFaker.City(),
				L12State = _states[Faker.NumberFaker.Number(0, 49)],
				L13Zip = Faker.LocationFaker.ZipCode(),
				L9StreetAddress = Faker.LocationFaker.StreetNumber() + " " + Faker.LocationFaker.StreetName(),
				PlanStartMonth = month
			};
		}
	}
}
