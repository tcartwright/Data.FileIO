// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ColumnInfoList.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace Data.FileIO.Core
{
	/// <summary>
	/// Class ColumnInfoList.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class ColumnInfoList<T> : List<ColumnInfo<T>>
	{
		/// <summary>
		/// Adds an item to the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <param name="headerName">Name of the header.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <param name="cellType">Type of the cell.</param>
		/// <param name="updateHeader">if set to <c>true</c> [update header].</param>
		/// <param name="formatCode">The format code.</param>
		public void Add(string headerName, Expression<Func<T, object>> valueFunction, ExcelCellType cellType = ExcelCellType.General, bool updateHeader = false, string formatCode = null)
		{
			ColumnInfo<T> item = new ColumnInfo<T>(headerName, valueFunction)
			{
				UpdateHeader = updateHeader,
				CellType = cellType,
				FormatCode = formatCode
			};
			this.Add(item);
		}
		/// <summary>
		/// Adds an item to the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <param name="columnName">Name of the column.</param>
		/// <param name="headerName">Name of the header.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <param name="cellType">Type of the cell.</param>
		/// <param name="updateHeader">if set to <c>true</c> [update header].</param>
		/// <param name="formatCode">The format code.</param>
		public void Add(string columnName, string headerName, Expression<Func<T, object>> valueFunction, ExcelCellType cellType = ExcelCellType.General, bool updateHeader = false, string formatCode = null)
		{
			ColumnInfo<T> item = new ColumnInfo<T>(columnName, headerName, valueFunction)
			{
				UpdateHeader = updateHeader,
				CellType = cellType,
				FormatCode = formatCode
			};
			this.Add(item);
		}
	}
}
