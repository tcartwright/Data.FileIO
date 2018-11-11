using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.Server;

namespace Data.DataImport
{



	/// <summary>
	/// Provides a mechanism for performing custom validation on each object when importing into a database.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="validationObject">The validation object.</param>
	/// <param name="errors">The errors.</param>
	public delegate void CustomValidation<T>(T validationObject, ref List<string> errors);

	/// <summary>
	/// Provides a way to use custom mapping for mapping a poco to the destination table type when importing data to a database.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="mapObject">The map object.</param>
	/// <param name="record">The record.</param>
	/// <param name="rowIndex">Index of the row.</param>
	/// <param name="errors">The errors.</param>
	public delegate void CustomMapSqlRecord<T>(T mapObject, SqlDataRecord record, int rowIndex, IEnumerable<string> errors);
}
