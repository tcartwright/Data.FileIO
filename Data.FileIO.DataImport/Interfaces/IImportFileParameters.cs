// ***********************************************************************
// Assembly         : Data.FileIO.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IImportFileParameters.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport.Interfaces;
using Data.FileIO;
using Data.FileIO.Common;
using Data.FileIO.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.FileIO.DataImport.Interfaces
{
	/// <summary>
	/// Interface IImportFileParameters
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public interface IImportFileParameters<T> : IImportEnumerableParameters<T>
	{
		/// <summary>
		/// Gets or sets the file parser.
		/// </summary>
		/// <value>The parser.</value>
		IFileParser Parser { get; set; }

		/// <summary>
		/// Gets or sets the mapper used for custom mapping of the dynamic row object to the poco.
		/// </summary>
		/// <value>The mapper.</value>
		FileValuesMap<T> FileValuesMapper { get; set; }
	}
}
