// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ColumnInfo.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Interfaces;

namespace Data.FileIO.Core
{
	/// <summary>
	/// Class ColumnInfo.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.FileIO.Core.ColumnInfoBase" />
	/// <seealso cref="Data.FileIO.Interfaces.IColumnInfo{T}" />
	public class ColumnInfo<T> : ColumnInfoBase, IColumnInfo<T>
	{
		/// <summary>
		/// Prevents a default instance of the <see cref="ColumnInfo{T}" /> class from being created.
		/// </summary>
		private ColumnInfo()
		{

		}
		/// <summary>
		/// Initializes a new instance of the <see cref="ColumnInfo{T}" /> class. The explicit column name is needed when customizing the sheet output or when overwriting an existing header.
		/// </summary>
		/// <param name="columnName">Name of the excel column minus the row number.</param>
		/// <param name="headerName">Name of the header in the template.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <exception cref="System.ArgumentOutOfRangeException">columnName</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
		public ColumnInfo(string columnName, string headerName, Expression<Func<T, object>> valueFunction)
			: this(headerName, valueFunction)
		{
			//TDC: It is very important that this ctor signature remains the same as it is used via reflection in the excel writer WriteDataToSheet method
			this.IsValidColumnName = Regex.IsMatch((columnName ?? String.Empty), "^[A-za-z]{1,3}$");
			//if they have provided us with a column name, validate it
			if (!String.IsNullOrWhiteSpace(columnName) && !this.IsValidColumnName)
			{
				throw new ArgumentOutOfRangeException("columnName");
			}
			this.ColumnName = columnName;
		}
		/// <summary>
		/// Initializes a new instance of the <see cref="ColumnInfo{T}" /> class. This overload only works if an existing header is already in the destination sheet that matches the supplied header name.
		/// </summary>
		/// <param name="headerName">Name of the header in the template.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <exception cref="System.ArgumentNullException">valueFunction</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
		public ColumnInfo(string headerName, Expression<Func<T, object>> valueFunction)
		{
			//if (String.IsNullOrWhiteSpace(headerName)) { throw new ArgumentNullException("headerName"); }
			if (valueFunction == null) { throw new ArgumentNullException("valueFunction"); }
			this.HeaderName = headerName;
			this.CleanHeaderName = FileIOUtilities.FixupHeader(this.HeaderName);
			this.ValueFunction = valueFunction.Compile();
			var type = FileIOUtilities.GetExpressionType<T>(valueFunction);
			var notNullType = Nullable.GetUnderlyingType(type);
			this.ValueFunctionType = notNullType ?? type;
			this.CellType = ExcelCellType.General;
		}

		/// <summary>
		/// Determines whether the specified column name is match for the current excel column or header. The column will take priority.
		/// </summary>
		/// <param name="columnName">Name of the column.</param>
		/// <param name="headerName">Name of the header.</param>
		/// <returns><c>true</c> if the specified column name is match; otherwise, <c>false</c>.</returns>
		public override bool IsMatch(string columnName, string headerName)
		{
			return (!String.IsNullOrWhiteSpace(this.ColumnName) && comparer.Equals(columnName, this.ColumnName))
				//only match the header name if the column name is empty to keep the map from matching two columns
				|| (
					String.IsNullOrWhiteSpace(this.ColumnName) 
					&& !String.IsNullOrWhiteSpace(this.HeaderName) 
					&& (comparer.Equals(this.HeaderName, headerName) || comparer.Equals(this.CleanHeaderName, headerName))
				);
		}

		/// <summary>
		/// Returns a <see cref="System.String" /> that represents this instance.
		/// </summary>
		/// <returns>A <see cref="System.String" /> that represents this instance.</returns>
		public override string ToString()
		{
			return String.Format("{0}  - {1}", this.ColumnName, this.HeaderName);
		}

		#region IColumnHeaderMap<T, K> Members

		/// <summary>
		/// Gets the value function. The value function will be used to provide the value for the cell. The row object being processed will be the parameter type.
		/// </summary>
		/// <value>The value function.</value>
		public Func<T, object> ValueFunction { get; private set; }

		/// <summary>
		/// Gets the type of the value function.
		/// </summary>
		/// <value>The type of the value function.</value>
		public Type ValueFunctionType { get; private set; }

		#endregion
	}
}
