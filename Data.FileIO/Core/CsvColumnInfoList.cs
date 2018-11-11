// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="CsvColumnInfoList.cs" company="Microsoft">
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
	/// Class CsvColumnInfoList.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class CsvColumnInfoList<T> : List<CsvColumnInfo<T>>
	{
		/// <summary>
		/// Adds an item to the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <param name="headerName">Name of the header.</param>
		/// <param name="valueFunction">The value function.</param>
		/// <param name="quotedFieldCharacter">The quoted field character.</param>
		public void Add(string headerName, Func<T, object> valueFunction, char? quotedFieldCharacter = null)
		{
			CsvColumnInfo<T> item = new CsvColumnInfo<T>(headerName, valueFunction, quotedFieldCharacter);
			this.Add(item);
		}
	}
}
