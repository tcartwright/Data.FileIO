// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ICsvColumnInfo.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace Data.FileIO.Interfaces
{
	/// <summary>
	/// Interface ICsvColumnInfo
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public interface ICsvColumnInfo<T>
	{
		/// <summary>
		/// Gets the name of the header. This header name will be used to match a column in the sheet if a column name is not supplied.
		/// </summary>
		/// <value>The name of the header.</value>
		string HeaderName { get; }

		/// <summary>
		/// Gets the value function. The value function will be used to provide the value for the cell. The row object being processed will be the parameter type.
		/// </summary>
		/// <value>The value function.</value>
		Func<T, object> ValueFunction { get; }

		/// <summary>
		/// Gets the quote.
		/// </summary>
		/// <value>The quoted field character.</value>
		string QuotedFieldCharacter { get; }

	}
}
