// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IFileParser.cs" company="Microsoft">
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
	/// Interface IFileParser
	/// </summary>
	public interface IFileParser
	{
		/// <summary>
		/// Parses the file.
		/// </summary>
		/// <returns>IEnumerable&lt;dynamic&gt;.</returns>
		IEnumerable<dynamic> ParseFile();

		/// <summary>
		/// Gets or sets the name of the file.
		/// </summary>
		/// <value>The name of the file.</value>
		string FileName { get; set; }

		/// <summary>
		/// Determines whether this instance has rows.
		/// </summary>
		/// <returns><c>true</c> if this instance has rows; otherwise, <c>false</c>.</returns>
		bool HasRows();

		/// <summary>
		/// Gets the row start.
		/// </summary>
		/// <value>The row start.</value>
		int RowStart { get; }

	}
}
