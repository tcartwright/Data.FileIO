// ***********************************************************************
// Assembly         : Data.FileIO.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ImportFileParameters.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport;
using Data.FileIO;
using Data.FileIO.Common;
using Data.FileIO.Core;
using Data.FileIO.DataImport.Interfaces;
using Data.FileIO.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.FileIO.DataImport
{
	/// <summary>
	/// Class ImportFileParameters.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.FileIO.DataImport.Interfaces.IImportFileParameters{T}" />
	public class ImportFileParameters<T> : ImportEnumerableParameters<T>, IImportFileParameters<T>
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="ImportFileParameters{T}"/> class.
		/// </summary>
		/// <param name="parser">The parser.</param>
		/// <param name="procedureName">Name of the procedure.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <param name="sqlMetadata">The SQL metadata.</param>
		/// <exception cref="System.ArgumentNullException">parser</exception>
		public ImportFileParameters(IFileParser parser, string procedureName, string tableTypeName = null, List<SqlMetaData> sqlMetadata = null)
			: base(null, procedureName, tableTypeName, sqlMetadata)
		{
			if (parser == null) { throw new ArgumentNullException("parser"); }
			this.Parser = parser;
		}

		#region IImportFileParameters<T> Members

		/// <summary>
		/// Gets or sets the parser.
		/// </summary>
		/// <value>The parser.</value>
		public IFileParser Parser { get; set; }

		/// <summary>
		/// Gets or sets the mapper.
		/// </summary>
		/// <value>The mapper.</value>
		public FileValuesMap<T> FileValuesMapper { get; set; }

		/// <summary>
		/// Checks the arguments to determine if they are valid.
		/// </summary>
		/// <exception cref="System.ArgumentNullException">Parser
		/// or
		/// SessionId
		/// or
		/// ProcedureName
		/// or
		/// Either the SqlMetadata or the TableName is required.</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2208:InstantiateArgumentExceptionsCorrectly")]
		public override void ArgumentsCheck()
		{
			if (this.Parser == null) { throw new ArgumentNullException("Parser"); }

			if (string.IsNullOrWhiteSpace(this.ProcedureName)) { throw new ArgumentNullException("ProcedureName"); }
			if (string.IsNullOrWhiteSpace(this.ProcedureParameterName)) { throw new ArgumentNullException("ProcedureParameterName"); }
			if (this.SqlMetadata == null && string.IsNullOrWhiteSpace(this.TableTypeName)) { throw new ArgumentNullException("Either the SqlMetadata or the TableName is required."); }
		}

		#endregion
	}
}
