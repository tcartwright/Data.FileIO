// ***********************************************************************
// Assembly         : Data.FileIO.DataImport
// Author           : tdcart
// Created          : 05-27-2016
//
// Last Modified By : tdcart
// Last Modified On : 05-27-2016
// ***********************************************************************
// <copyright file="IFileDataImportRepository.cs" company="Tim Cartwright">
//     Copyright © Tim Cartwright 2016
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Data;
using System.Data.SqlClient;
namespace Data.FileIO.DataImport.Interfaces
{
	/// <summary>
	/// Interface IFileDataImportRepository
	/// </summary>
	public interface IFileDataImportRepository
	{
		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportFile<T>(SqlConnection connection, IImportFileParameters<T> parameters);
		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportFile<T>(SqlTransaction transaction, IImportFileParameters<T> parameters);
		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		void ImportFile<T>(string connectionString, IImportFileParameters<T> parameters);
		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>System.Data.IDataReader.</returns>
		IDataReader ImportFileReader<T>(SqlConnection connection, IImportFileParameters<T> parameters);
		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>System.Data.IDataReader.</returns>
		IDataReader ImportFileReader<T>(SqlTransaction transaction, IImportFileParameters<T> parameters);
		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>System.Data.IDataReader.</returns>
		IDataReader ImportFileReader<T>(string connectionString, IImportFileParameters<T> parameters);
	}
}
