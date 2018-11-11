// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IExcelFileParser.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************

using System.Collections.Generic;

namespace Data.FileIO.Interfaces
{
	/// <summary>
	/// Interface IExcelFileParser
	/// </summary>
	/// <seealso cref="Data.FileIO.Interfaces.IFileParser" />
	public interface IExcelFileParser : IFileParser
	{
		/// <summary>
		/// Gets or sets the name of the sheet.
		/// </summary>
		/// <value>The name of the sheet.</value>
		string SheetName { get; set; }
		/// <summary>
		/// Gets or sets the header row.
		/// </summary>
		/// <value>The header row.</value>
		int HeaderRow { get; set; }
		/// <summary>
		/// Gets or sets the data row start.
		/// </summary>
		/// <value>The data row start.</value>
		int DataRowStart { get; set; }
		/// <summary>
		/// Gets or sets the data row end.
		/// </summary>
		/// <value>The data row end.</value>
		int DataRowEnd { get; set; }
		/// <summary>
		/// The 1st column in the sheet that is valid for parsing, If null, then A will be assumed. Useful when a sheet contains multiple data sets.
		/// </summary>
		/// <value>The start column key.</value>
		string StartColumnKey { get; set; }

		/// <summary>
		/// The very last column that should be parsed. If null, the entire sheet will be parsed. Useful when a sheet contains multiple data sets.
		/// </summary>
		/// <value>The end column key.</value>
		string EndColumnKey { get; set; }

		/// <summary>
		/// Gets the headers for the specified worksheet.
		/// </summary>
		/// <param name="removeInvalidChars">if set to <c>true</c> [remove invalid chars].</param>
		/// <returns>IDictionary&lt;System.String, System.String&gt;.</returns>
		IDictionary<string, string> GetHeaders(bool removeInvalidChars = false);
	}
}
