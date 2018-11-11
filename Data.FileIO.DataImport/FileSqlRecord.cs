// ***********************************************************************
// Assembly         : Data.FileIO.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-27-2016
// ***********************************************************************
// <copyright file="FileSqlRecord.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;
using Data.DataImport.SqlRecordImporters;
using Data.FileIO.DataImport.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.FileIO.DataImport
{
	/// <summary>
	/// Class FileSqlRecord.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class FileSqlRecord<T> : EnumerableSqlRecord<T>
	{

		/// <summary>
		/// Initializes a new instance of the <see cref="FileSqlRecord{T}"/> class.
		/// </summary>
		/// <param name="parameters">The parameters.</param>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0")]
		public FileSqlRecord(IImportFileParameters<T> parameters)
			: base(parameters, parameters.Parser.HasRows())
		{
			this._data = parameters.Parser.ParseFile();
			this._rowIndex = Math.Max(parameters.Parser.RowStart - 1, 1);
		}
	}
}
