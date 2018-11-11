// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="CsvColumnInfo.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Data.FileIO.Interfaces;

namespace Data.FileIO.Core
{
	/// <summary>
	/// Class CsvColumnInfo.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.FileIO.Interfaces.ICsvColumnInfo{T}" />
	public class CsvColumnInfo<T> : ICsvColumnInfo<T>
	{
		/// <summary>
		/// Prevents a default instance of the <see cref="CsvColumnInfo{T}"/> class from being created.
		/// </summary>
		private CsvColumnInfo() { }
		/// <summary>
		/// Initializes a new instance of the <see cref="CsvColumnInfo{T}"/> class.
		/// </summary>
		/// <param name="headerName">Name of the header.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <param name="quotedFieldCharacter">The quoted field character.</param>
		public CsvColumnInfo(string headerName, Func<T, object> valueFunction, char? quotedFieldCharacter = null)
		{
			this.HeaderName = headerName;
			this.ValueFunction = valueFunction;
			if (quotedFieldCharacter.HasValue) { this.QuotedFieldCharacter = quotedFieldCharacter.Value.ToString(); }
		}

		#region ICsvColumnInfo<T> Members

		/// <summary>
		/// Gets the name of the header. This header name will be used to match a column in the sheet if a column name is not supplied.
		/// </summary>
		/// <value>The name of the header.</value>
		public string HeaderName { get; private set; }

		/// <summary>
		/// Gets the value function. The value function will be used to provide the value for the cell. The row object being processed will be the parameter type.
		/// </summary>
		/// <value>The value function.</value>
		public Func<T, object> ValueFunction { get; private set; }

		/// <summary>
		/// Gets the quote.
		/// </summary>
		/// <value>The quoted field character.</value>
		public string QuotedFieldCharacter { get; private set; }
		#endregion

		/// <summary>
		/// Returns a <see cref="System.String" /> that represents this instance.
		/// </summary>
		/// <returns>A <see cref="System.String" /> that represents this instance.</returns>
		public override string ToString()
		{
			return this.HeaderName;
		}
	}
}
