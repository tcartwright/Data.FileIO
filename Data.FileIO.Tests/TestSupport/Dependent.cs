using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport.Interfaces;
using Data.FileIO.Common.Interfaces;

namespace Data.FileIOTests.TestSupport
{
	public class Dependent : IFileRowMapper
	{
		public string L2EmployeeSSN { get; set; }
		public string aName { get; set; }
		public string bSSN { get; set; }
		public DateTime cDOB { get; set; }

		#region IFileRowMapper Members

		public void MapValues(int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors)
		{
			this.L2EmployeeSSN = row.L2EmployeeSSN;
			this.aName = row.aName;
			this.bSSN = row.bSSN;
			this.cDOB = validator.GetRowValue<DateTime>(row, "cDob", ref errors);
		}

		#endregion
	}
}
