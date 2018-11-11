// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IColumnInfoBase.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Data.FileIO.Core;

namespace Data.FileIO.Interfaces
{
	/// <summary>
	/// Interface IColumnInfoBase
	/// </summary>
	public interface IColumnInfoBase
	{
		/// <summary>
		/// Gets the name of the excel column. A column name is absolute. The header, and data will be written to the exact column if supplied.
		/// </summary>
		/// <value>The name of the column.</value>
		string ColumnName { get; }
		/// <summary>
		/// Gets the name of the header. This header name will be used to match a column in the sheet if a column name is not supplied.
		/// </summary>
		/// <value>The name of the header.</value>
		string HeaderName { get; }

		/// <summary>
		/// Gets the value function string. This is only used when using an <see cref="System.Data.IDataReader" /> to write the sheet without a strong type.
		/// </summary>
		/// <value>The value function string.</value>
		string ValueFunctionString { get; }

		/// <summary>
		/// Gets a value indicating whether this instance is valid excel column name. Valid names are from 1 to 3 alpha characters.
		/// </summary>
		/// <value><c>true</c> if this instance is valid column name; otherwise, <c>false</c>.</value>
		bool IsValidColumnName { get; }

		/// <summary>
		/// Gets or sets the type of formatting to apply to the cell.
		/// </summary>
		/// <value>The type of the cell.</value>
		ExcelCellType CellType { get; set; }

		/// <summary>
		/// Gets or sets the format code. Allows for custom formatting to be applied to the cell outside of the default formatting. If left empty the default formatting for the <seealso cref="ExcelCellType" /> will be used.
		/// </summary>
		/// <value>The format code.</value>
		string FormatCode { get; set; }

		/// <summary>
		/// Gets a value indicating whether [update header].
		/// </summary>
		/// <value><c>true</c> if [update header]; otherwise, <c>false</c>.</value>
		bool UpdateHeader { get; set; }

		/// <summary>
		/// Determines whether the specified column name is match for the current excel column or header. The column will take priority.
		/// </summary>
		/// <param name="columnName">Name of the column.</param>
		/// <param name="headerName">Name of the header.</param>
		/// <returns><c>true</c> if the specified column name is match; otherwise, <c>false</c>.</returns>
		bool IsMatch(string columnName, string headerName);

	}
}
