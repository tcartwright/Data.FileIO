// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ImportEnumerableParameters.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport.Interfaces;
using Data.FileIO.Common;
using Data.FileIO.Common.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.DataImport
{
	/// <summary>
	/// Class ImportEnumerableParameters.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.DataImport.Interfaces.IImportEnumerableParameters{T}" />
	public class ImportEnumerableParameters<T> : IImportEnumerableParameters<T>
	{
		private bool _requireProcedureName = true;

		/// <summary>
		/// Initializes a new instance of the <see cref="ImportEnumerableParameters{T}" /> class.
		/// </summary>
		/// <param name="data">The data.</param>
		/// <param name="sqlMetadata">The SQL metadata.</param>
		/// <exception cref="ArgumentException">sqlMetadata</exception>
		public ImportEnumerableParameters(IEnumerable<dynamic> data, IList<SqlMetaData> sqlMetadata)
		{
			if (sqlMetadata == null) { throw new ArgumentNullException("sqlMetadata"); }
			SetupDefaults();

			_requireProcedureName = false;
			this.Data = data;
			this.SqlMetadata.AddRange(sqlMetadata);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ImportEnumerableParameters{T}"/> class.
		/// </summary>
		/// <param name="data">The data.</param>
		/// <param name="procedureName">Name of the procedure.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <param name="sqlMetadata">The SQL metadata.</param>
		public ImportEnumerableParameters(IEnumerable<dynamic> data, string procedureName, string tableTypeName = null, IList<SqlMetaData> sqlMetadata = null)
		{
			_requireProcedureName = true;
			if (String.IsNullOrWhiteSpace(procedureName)) { throw new ArgumentNullException("procedureName"); }
			//data, tabletypename, and sqlmetadata can all be set later
			SetupDefaults();

			this.Data = data;
			this.ProcedureName = procedureName;
			this.TableTypeName = tableTypeName;
			if (sqlMetadata != null) { this.SqlMetadata.AddRange(sqlMetadata); }

		}

		private void SetupDefaults()
		{
			//set up the defaults
			this.ProcedureParameterName = "@data";
			this.Validator = new ObjectValidator();
			this.ErrorColumnName = "Errors";
			this.ExtraParameters = new List<SqlParameter>();
			this.SqlMetadata = new List<SqlMetaData>();
		}

		#region IImportEnumerableParameters<T> Members

		/// <summary>
		/// Gets or sets the data.
		/// </summary>
		/// <value>The data.</value>
		public IEnumerable<dynamic> Data { get; set; }

		/// <summary>
		/// Gets or sets the name of the table type.
		/// </summary>
		/// <value>The name of the table type.</value>
		public string TableTypeName { get; set; }

		/// <summary>
		/// Gets or sets the name of the procedure.
		/// </summary>
		/// <value>The name of the procedure.</value>
		public string ProcedureName { get; set; }

		/// <summary>
		/// Gets or sets the name of the procedure parameter.
		/// </summary>
		/// <value>The name of the procedure parameter.</value>
		public string ProcedureParameterName { get; set; }

		/// <summary>
		/// Gets or sets the custom validator.
		/// </summary>
		/// <value>The custom validator.</value>
		public CustomValidation<T> CustomValidator { get; set; }

		/// <summary>
		/// Gets or sets the custom SQL mapper.
		/// </summary>
		/// <value>The custom SQL mapper.</value>
		public CustomMapSqlRecord<T> CustomSqlMapper { get; set; }
		/// <summary>
		/// Gets or sets the SQL metadata.
		/// </summary>
		/// <value>The SQL metadata.</value>
		public List<SqlMetaData> SqlMetadata { get; internal set; }
		/// <summary>
		/// Gets or sets the extra parameters for the procedure call. Do not add the table data parameter to this collection.
		/// </summary>
		/// <value>The extra parameters.</value>
		public List<SqlParameter> ExtraParameters { get; internal set; }
		/// <summary>
		/// Gets or sets the validator.
		/// </summary>
		/// <value>The validator.</value>
		public IObjectValidator Validator { get; set; }

		/// <summary>
		/// Gets or sets the name of the error column.
		/// </summary>
		/// <value>The name of the error column.</value>
		public string ErrorColumnName { get; set; }

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
		public virtual void ArgumentsCheck()
		{
			if (_requireProcedureName && string.IsNullOrWhiteSpace(this.ProcedureName)) { throw new ArgumentNullException("ProcedureName"); }
			if (string.IsNullOrWhiteSpace(this.ProcedureParameterName)) { throw new ArgumentNullException("ProcedureParameterName"); }
			if (this.SqlMetadata == null && string.IsNullOrWhiteSpace(this.TableTypeName)) { throw new ArgumentNullException("Either the SqlMetadata or the TableName is required."); }
		}

		#endregion
	}
}
