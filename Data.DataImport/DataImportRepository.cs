// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="DataImportRepository.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Data.DataImport.Interfaces;
using Data.DataImport.SqlRecordImporters;
using Microsoft.SqlServer.Server;

namespace Data.DataImport
{
	/// <summary>
	/// Class DataImportRepository.
	/// </summary>
	/// <seealso cref="Data.DataImport.Interfaces.IDataImportRepository" />
	public class DataImportRepository : IDataImportRepository
	{
		private int _commandTimeout;
		/// <summary>
		/// Initializes a new instance of the <see cref="DataImportRepository"/> class.
		/// </summary>
		public DataImportRepository()
			: this(30)
		{

		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DataImportRepository"/> class.
		/// </summary>
		/// <param name="commandTimeout">The command timeout.</param>
		public DataImportRepository(int commandTimeout)
		{
			_commandTimeout = Math.Max(commandTimeout, 5);
		}
		#region reader imports
		/// <summary>
		/// Imports the enumerable reader.
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
		public virtual IDataReader ImportEnumerableReader<T>(string connectionString, IImportEnumerableParameters<T> parameters)
		{
			if (String.IsNullOrWhiteSpace(connectionString)) { throw new ArgumentNullException("connectionString"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				return this.ImportData<T>(null, connection, parameters, new EnumerableSqlRecord<T>(parameters), true);
			}
		}

		/// <summary>
		/// Imports the enumerable reader.
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
		public virtual IDataReader ImportEnumerableReader<T>(SqlTransaction transaction, IImportEnumerableParameters<T> parameters)
		{
			if (transaction == null) { throw new ArgumentNullException("transaction"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			return this.ImportData<T>(transaction, transaction.Connection, parameters, new EnumerableSqlRecord<T>(parameters), true);
		}

		/// <summary>
		/// Imports the enumerable reader.
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
		public virtual IDataReader ImportEnumerableReader<T>(SqlConnection connection, IImportEnumerableParameters<T> parameters)
		{
			if (connection == null) { throw new ArgumentNullException("connection"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			return this.ImportData<T>(null, connection, parameters, new EnumerableSqlRecord<T>(parameters), true);
		}
		#endregion

		#region enumerable import
		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">connectionString
		/// or
		/// parameters</exception>
		public virtual void ImportEnumerable<T>(string connectionString, IImportEnumerableParameters<T> parameters)
		{
			if (String.IsNullOrWhiteSpace(connectionString)) { throw new ArgumentNullException("connectionString"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				this.ImportData<T>(null, connection, parameters, new EnumerableSqlRecord<T>(parameters));
			}
		}
		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">transaction
		/// or
		/// parameters</exception>
		public virtual void ImportEnumerable<T>(SqlTransaction transaction, IImportEnumerableParameters<T> parameters)
		{
			if (transaction == null) { throw new ArgumentNullException("transaction"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			this.ImportData<T>(transaction, transaction.Connection, parameters, new EnumerableSqlRecord<T>(parameters));
		}

		/// <summary>
		/// Imports the enumerable.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <exception cref="System.ArgumentNullException">connection
		/// or
		/// parameters</exception>
		public virtual void ImportEnumerable<T>(SqlConnection connection, IImportEnumerableParameters<T> parameters)
		{
			if (connection == null) { throw new ArgumentNullException("connection"); }
			if (parameters == null) { throw new ArgumentNullException("parameters"); }

			this.ImportData<T>(null, connection, parameters, new EnumerableSqlRecord<T>(parameters));
		}

		#endregion

		#region Misc
		/// <summary>
		/// Imports the data.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="transaction">The transaction.</param>
		/// <param name="connection">The connection.</param>
		/// <param name="parameters">The parameters.</param>
		/// <param name="data">The data.</param>
		/// <param name="returnReader">if set to <c>true</c> [return reader].</param>
		/// <returns>IDataReader.</returns>
		/// <exception cref="System.ArgumentNullException">parameters
		/// or
		/// data
		/// or
		/// connection</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")]
		protected virtual IDataReader ImportData<T>(SqlTransaction transaction, SqlConnection connection, IImportEnumerableParameters<T> parameters, IEnumerableSqlRecord data, bool returnReader = false)
		{
			if (parameters == null) { throw new ArgumentNullException("parameters"); }
			if (data == null) { throw new ArgumentNullException("data"); }
			if (connection == null) { throw new ArgumentNullException("connection"); }

			//check the parameters to test if they are valid, if not, an exception will be raised
			parameters.ArgumentsCheck();

			using (SqlCommand command = new SqlCommand(parameters.ProcedureName, connection))
			{
				if (transaction != null) { command.Transaction = transaction; }
				command.CommandTimeout = _commandTimeout;
				command.CommandType = CommandType.StoredProcedure;
				//only add the table parameter IF the data has values. Otherwise sql server will throw an exception
				if (data.HasRows())
				{
					//only get the meta data if we have rows and the metadata is empty, as it has no impact otherwise
					if (data.SqlMetadata == null || data.SqlMetadata.Count == 0)
					{
						data.SqlMetadata.AddRange(this.GetSqlMetadata(transaction, connection, parameters.TableTypeName));
					}
					command.Parameters.Add(new SqlParameter(parameters.ProcedureParameterName, SqlDbType.Structured) { Value = data });
				}

				//add the extra parameters into the command
				if (parameters.ExtraParameters != null)
				{
					foreach (var param in parameters.ExtraParameters)
					{
						command.Parameters.Add(param);
					}
				}

				if (returnReader)
				{
					return command.ExecuteReader();
				}
				else
				{
					command.ExecuteNonQuery();
					return null;
				}
			}

		}

		/// <summary>
		/// Gets the SQL metadata.
		/// </summary>
		/// <param name="connectionString">The connection string.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <returns>List&lt;SqlMetaData&gt;.</returns>
		public virtual List<SqlMetaData> GetSqlMetadata(string connectionString, string tableTypeName)
		{
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				connection.Open();
				return this.GetSqlMetadata(null, connection, tableTypeName);
			}
		}

		/// <summary>
		/// Gets the SQL metadata.
		/// </summary>
		/// <param name="transaction">The transaction.</param>
		/// <param name="connection">The connection.</param>
		/// <param name="tableTypeName">Name of the table type.</param>
		/// <returns>List&lt;SqlMetaData&gt;.</returns>
		/// <exception cref="System.ArgumentNullException">tableTypeName</exception>
		/// <exception cref="System.InvalidOperationException">Could not build sql meta data. Please check that the table type exists.</exception>
		/// <exception cref="ArgumentNullException">tableTypeName</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")]
		public virtual List<SqlMetaData> GetSqlMetadata(SqlTransaction transaction, SqlConnection connection, string tableTypeName)
		{
			if (String.IsNullOrWhiteSpace(tableTypeName)) { throw new ArgumentNullException("tableTypeName"); }
			List<SqlMetaData> metas = new List<SqlMetaData>();
			StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;

			using (SqlCommand command = new SqlCommand(Data.DataImport.Properties.Resources.GetMetaData, connection))
			{
				command.CommandType = CommandType.Text;

				SqlParameter sqlParameter = new SqlParameter("@tableTypeName", SqlDbType.NVarChar, 512) { Value = tableTypeName };
				command.Parameters.Add(sqlParameter);
				if (transaction != null)
				{
					command.Transaction = transaction;
				}
				using (SqlDataReader reader = command.ExecuteReader())
				{
					while (reader.Read())
					{
						bool isIdentity = reader.GetBoolean(reader.GetOrdinal("IsIdentity"));
						if (!isIdentity)
						{
							string typeName = reader.GetString(reader.GetOrdinal("TypeName")).ToLower();
							string name = reader.GetString(reader.GetOrdinal("name"));
							int maxLength = reader.GetInt16(reader.GetOrdinal("max_length"));

							SqlMetaData meta = null;
							SqlDbType dbType = SqlDbType.Variant;
							if (comparer.Equals(typeName, "hierarchyid"))
							{
								dbType = SqlDbType.NVarChar;
								maxLength = Math.Min(maxLength * 2, 4000);
							}
							else if (comparer.Equals(typeName, "numeric")) //not parseable
							{
								dbType = SqlDbType.Decimal;
							}
							else if (!Enum.TryParse<SqlDbType>(typeName, true, out dbType))
							{
								//many types are not mappable: geometry, geography, matrix. treat anything not mappable as sqlvariant :| Hopefully that works.
								dbType = SqlDbType.Variant;
							}

							switch (dbType)
							{
								case SqlDbType.Char:
								case SqlDbType.NChar:
								case SqlDbType.VarChar:
								case SqlDbType.NVarChar:
								case SqlDbType.Binary:
								case SqlDbType.VarBinary:
									meta = new SqlMetaData(name, dbType, maxLength);
									break;
								case SqlDbType.Decimal:
									meta = new SqlMetaData(name, dbType, reader.GetByte(reader.GetOrdinal("precision")), reader.GetByte(reader.GetOrdinal("scale")));
									break;
								default:
									meta = new SqlMetaData(name, dbType);
									break;
							}
							metas.Add(meta);
						}
					}
				}
			}

			if (metas.Count == 0)
			{
				throw new InvalidOperationException("Could not build sql meta data. Please check that the table type exists.");
			}

			return metas;
		}

		#endregion Misc
	}

}
