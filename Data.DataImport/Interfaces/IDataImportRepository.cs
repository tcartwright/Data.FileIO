// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IDataImportRepository.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SqlServer.Server;

namespace Data.DataImport.Interfaces
{
	/// <summary>
	/// Interface IDataImportRepository
	/// </summary>
	public interface IDataImportRepository
	{

		/// <summary>
		/// Imports the enumerable and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		IDataReader ImportEnumerableReader<T>(string connectionString, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Imports the enumerable and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		IDataReader ImportEnumerableReader<T>(SqlConnection connection, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Imports the enumerable and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		IDataReader ImportEnumerableReader<T>(SqlTransaction transaction, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportEnumerable<T>(string connectionString, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportEnumerable<T>(SqlTransaction transaction, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportEnumerable<T>(SqlConnection connection, IImportEnumerableParameters<T> parameters);

		/// <summary>
		/// Gets the SQL metadata.
		/// </summary>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <returns>List&lt;SqlMetaData&gt;.</returns>
		List<SqlMetaData> GetSqlMetadata(string connectionString, string tableTypeName);
		/// <summary>
		/// Gets the SQL metadata.
		/// </summary>
		/// <param name="transaction">The transaction.</param>
		/// <param name="connection">The connection.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <returns>List&lt;SqlMetaData&gt;.</returns>
		List<SqlMetaData> GetSqlMetadata(SqlTransaction transaction, SqlConnection connection, string tableTypeName);
	}
}
