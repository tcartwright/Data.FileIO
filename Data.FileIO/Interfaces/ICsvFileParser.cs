// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ICsvFileParser.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Data.FileIO.Interfaces
{
	/// <summary>
	/// Interface ICsvFileParser
	/// </summary>
	/// <seealso cref="Data.FileIO.Interfaces.IFileParser" />
	public interface ICsvFileParser : IFileParser
	{
		/// <summary>
		/// Gets or sets the fixed widths.
		/// </summary>
		/// <value>The fixed widths.</value>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
		int[] FixedWidths { get; set; }

		/// <summary>
		/// Gets or sets the delimiters.
		/// </summary>
		/// <value>The delimiters.</value>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays")]
		string[] Delimiters { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether [first row contains headers].
		/// </summary>
		/// <value><c>true</c> if [first row contains headers]; otherwise, <c>false</c>.</value>
		bool FirstRowContainsHeaders { get; set; }

		/// <summary>
		/// Gets or sets the rows to skip.
		/// </summary>
		/// <value>The rows to skip.</value>
		int RowsToSkip { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether this instance has quoted fields.
		/// </summary>
		/// <value><c>true</c> if this instance has quoted fields; otherwise, <c>false</c>.</value>
		bool HasQuotedFields { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether [trim white space].
		/// </summary>
		/// <value><c>true</c> if [trim white space]; otherwise, <c>false</c>.</value>
		bool TrimWhiteSpace { get; set; }
	}
}
