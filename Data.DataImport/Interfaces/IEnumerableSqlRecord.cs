// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-30-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-30-2016
// ***********************************************************************
// <copyright file="IEnumerableSqlRecord.cs" company="Tim Cartwright">
//     Copyright © Tim Cartwright 2016
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using Microsoft.SqlServer.Server;
namespace Data.DataImport.Interfaces
{
	/// <summary>
	/// Interface IEnumerableSqlRecord
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public interface IEnumerableSqlRecord : IEnumerable<SqlDataRecord> 
	{
		/// <summary>
		/// Gets the SQL metadata.
		/// </summary>
		/// <value>The SQL metadata.</value>
		List<SqlMetaData> SqlMetadata { get; }

		/// <summary>
		/// Determines whether this instance has rows.
		/// </summary>
		/// <returns><c>true</c> if this instance has rows; otherwise, <c>false</c>.</returns>
		bool HasRows();
	}
}
