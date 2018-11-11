// ***********************************************************************
// Assembly         : Data.FileIO.DataImport
// Author           : tdcart
// Created          : 05-27-2016
//
// Last Modified By : tdcart
// Last Modified On : 05-27-2016
// ***********************************************************************
// <copyright file="FileDataImportRepository.cs" company="Tim Cartwright">
//     Copyright © Tim Cartwright 2016
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using Data.DataImport;
using Data.FileIO.DataImport.Interfaces;

namespace Data.FileIO.DataImport
{
	/// <summary>
	/// Class FileDataImportRepository.
	/// </summary>
	/// <seealso cref="Data.DataImport.DataImportRepository" />
	public class FileDataImportRepository : DataImportRepository, IFileDataImportRepository
	{

		#region Custom file reader

		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// connectionString
		/// or
		/// parameters
		/// </exception>
		public virtual IDataReader ImportFileReader<T>(string connectionString, IImportFileParameters<T> parameters)
		{
			if (String.IsNullOrWhiteSpace(connectionString)) { throw new ArgumentNullException("connectionString"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				return this.ImportData<T>(null, connection, parameters, new FileSqlRecord<T>(parameters), true);
			}
		}

		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// transaction
		/// or
		/// parameters
		/// </exception>
		public virtual IDataReader ImportFileReader<T>(SqlTransaction transaction, IImportFileParameters<T> parameters)
		{
			if (transaction == null) { throw new ArgumentNullException("transaction"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			return this.ImportData<T>(transaction, transaction.Connection, parameters, new FileSqlRecord<T>(parameters), true);
		}
		/// <summary>
		/// Imports the file and returns a reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <returns>IDataReader.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// connection
		/// or
		/// parameters
		/// </exception>
		public virtual IDataReader ImportFileReader<T>(SqlConnection connection, IImportFileParameters<T> parameters)
		{
			if (connection == null) { throw new ArgumentNullException("connection"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			return this.ImportData<T>(null, connection, parameters, new FileSqlRecord<T>(parameters), true);
		}

		#endregion Custom file

		#region Custom file

		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">
		/// connectionString
		/// or
		/// parameters
		/// </exception>
		public virtual void ImportFile<T>(string connectionString, IImportFileParameters<T> parameters)
		{
			if (String.IsNullOrWhiteSpace(connectionString)) { throw new ArgumentNullException("connectionString"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				this.ImportData<T>(null, connection, parameters, new FileSqlRecord<T>(parameters));
			}
		}

		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">
		/// transaction
		/// or
		/// parameters
		/// </exception>
		public virtual void ImportFile<T>(SqlTransaction transaction, IImportFileParameters<T> parameters)
		{
			if (transaction == null) { throw new ArgumentNullException("transaction"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			this.ImportData<T>(transaction, transaction.Connection, parameters, new FileSqlRecord<T>(parameters));
		}
		/// <summary>
		/// Imports the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">
		/// connection
		/// or
		/// parameters
		/// </exception>
		public virtual void ImportFile<T>(SqlConnection connection, IImportFileParameters<T> parameters)
		{
			if (connection == null) { throw new ArgumentNullException("connection"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			this.ImportData<T>(null, connection, parameters, new FileSqlRecord<T>(parameters));
		}

		#endregion Custom file

	}
}
