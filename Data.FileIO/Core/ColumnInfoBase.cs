// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ColumnInfoBase.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Interfaces;

namespace Data.FileIO.Core
{
	/// <summary>
	/// Class ColumnInfoBase.
	/// </summary>
	/// <seealso cref="Data.FileIO.Interfaces.IColumnInfoBase" />
	public class ColumnInfoBase : IColumnInfoBase
	{
		/// <summary>
		/// The comparer
		/// </summary>
		protected StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;
		/// <summary>
		/// Prevents a default instance of the <see cref="ColumnInfoBase" /> class from being created.
		/// </summary>
		internal ColumnInfoBase()
		{

		}
		/// <summary>
		/// Initializes a new instance of the <see cref="ColumnInfoBase" /> class. The explicit column name is needed when customizing the sheet output or when overwriting an existing header.
		/// </summary>
		/// <param name="columnName">Name of the excel column minus the row number.</param>
		/// <param name="headerName">Name of the header in the template.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <exception cref="System.ArgumentOutOfRangeException">columnName</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
		public ColumnInfoBase(string columnName, string headerName, string valueFunction)
			: this(headerName, valueFunction)
		{
			this.IsValidColumnName = Regex.IsMatch((columnName ?? String.Empty), "^[A-za-z]{1,3}$");
			//if they have provided us with a column name, validate it
			if (!String.IsNullOrWhiteSpace(columnName) && !this.IsValidColumnName)
			{
				throw new ArgumentOutOfRangeException("columnName");
			}
			this.ColumnName = columnName;
		}
		/// <summary>
		/// Initializes a new instance of the <see cref="ColumnInfoBase" /> class. This overload only works if an existing header is already in the destination sheet that matches the supplied header name.
		/// </summary>
		/// <param name="headerName">Name of the header in the template.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <exception cref="System.ArgumentNullException">headerName</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
		public ColumnInfoBase(string headerName, string valueFunction)
		{
			//if (String.IsNullOrWhiteSpace(headerName)) { throw new ArgumentNullException("headerName"); }
			if (valueFunction == null) { throw new ArgumentNullException("valueFunction"); }
			this.HeaderName = headerName;
			this.CleanHeaderName = FileIOUtilities.FixupHeader(this.HeaderName);
			this.ValueFunctionString = valueFunction;
			this.CellType = ExcelCellType.General; 
		}

		/// <summary>
		/// Gets the name of the excel column. A column name is absolute. The header, and data will be written to the exact column if supplied.
		/// </summary>
		/// <value>The name of the column.</value>
		public virtual string ColumnName { get; internal set; }
		/// <summary>
		/// Gets the name of the header. This header name will be used to match a column in the sheet if a column name is not supplied.
		/// </summary>
		/// <value>The name of the header.</value>
		public virtual string HeaderName { get; internal set; }

		/// <summary>
		/// Gets or sets the name of the clean header.
		/// </summary>
		/// <value>The name of the clean header.</value>
		protected virtual string CleanHeaderName { get; set; }

		/// <summary>
		/// Gets the value function string. This is only used when using an <see cref="System.Data.IDataReader" /> to write the sheet without a strong type.
		/// </summary>
		/// <value>The value function string.</value>
		public virtual string ValueFunctionString { get; internal set; }

		/// <summary>
		/// Gets a value indicating whether this instance is valid excel column name. Valid names are from 1 to 3 alpha characters.
		/// </summary>
		/// <value><c>true</c> if this instance is valid column name; otherwise, <c>false</c>.</value>
		public virtual bool IsValidColumnName { get; internal set; }

		/// <summary>
		/// Gets a value indicating whether [update header].
		/// </summary>
		/// <value><c>true</c> if [update header]; otherwise, <c>false</c>.</value>
		public virtual bool UpdateHeader { get; set; }

		/// <summary>
		/// Gets or sets the type of formatting to apply to the cell.
		/// </summary>
		/// <value>The type of the cell.</value>
		public virtual ExcelCellType CellType { get; set; }

		/// <summary>
		/// Gets or sets the format code. Allows for custom formatting to be applied to the cell outside of the default formatting. If left empty the default formatting for the <seealso cref="ExcelCellType" /> will be used.
		/// </summary>
		/// <value>The format code.</value>
		public virtual string FormatCode { get; set; }

		/// <summary>
		/// Determines whether the specified column name is match for the current excel column or header. The column will take priority.
		/// </summary>
		/// <param name="columnName">Name of the column.</param>
		/// <param name="headerName">Name of the header.</param>
		/// <returns><c>true</c> if the specified column name is match; otherwise, <c>false</c>.</returns>
		public virtual bool IsMatch(string columnName, string headerName)
		{
			return false;
		}
	}
}
