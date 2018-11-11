// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-27-2016
// ***********************************************************************
// <copyright file="IImportEnumerableParameters.cs" company="Microsoft">
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
using Data.FileIO.Common.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.DataImport.Interfaces
{
	/// <summary>
	/// Interface IImportFileParameters
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public interface IImportEnumerableParameters<T>
	{
		/// <summary>
		/// Gets or sets the data.
		/// </summary>
		/// <value>The data.</value>
		IEnumerable<dynamic> Data { get; set; }

		/// <summary>
		/// Gets or sets the name of the table type.
		/// </summary>
		/// <value>The name of the table type.</value>
		string TableTypeName { get; set; }

		/// <summary>
		/// Gets or sets the name of the procedure.
		/// </summary>
		/// <value>The name of the procedure.</value>
		string ProcedureName { get; set; }
		/// <summary>
		/// Gets or sets the name of the procedure parameter.
		/// </summary>
		/// <value>The name of the procedure parameter.</value>
		string ProcedureParameterName { get; set; }
		/// <summary>
		/// Gets or sets the custom validator.
		/// </summary>
		/// <value>The custom validator.</value>
		CustomValidation<T> CustomValidator { get; set; }

		/// <summary>
		/// Gets or sets the custom SQL mapper.
		/// </summary>
		/// <value>The custom SQL mapper.</value>
		CustomMapSqlRecord<T> CustomSqlMapper { get; set; }

		/// <summary>
		/// Gets or sets the SQL metadata.
		/// </summary>
		/// <value>The SQL metadata.</value>
		List<SqlMetaData> SqlMetadata { get; }
		
		/// <summary>
		/// Gets or sets the extra parameters for the procedure call. Do not add the table data parameter to this collection.
		/// </summary>
		/// <value>The extra parameters.</value>
		List<SqlParameter> ExtraParameters { get; }

		/// <summary>
		/// Gets or sets the validator.
		/// </summary>
		/// <value>The validator.</value>
		IObjectValidator Validator { get; set; }
		/// <summary>
		/// Gets or sets the name of the error column.
		/// </summary>
		/// <value>The name of the error column.</value>
		string ErrorColumnName { get; set; }
		/// <summary>
		/// Checks the arguments to determine if they are valid.
		/// </summary>
		void ArgumentsCheck();
	}
}
