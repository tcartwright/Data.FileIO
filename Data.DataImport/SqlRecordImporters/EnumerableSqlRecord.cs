// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-27-2016
// ***********************************************************************
// <copyright file="EnumerableSqlRecord.cs" company="Microsoft">
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
using Data.DataImport.Interfaces;
using Data.FileIO;
using Data.FileIO.Common;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Microsoft.SqlServer.Server;

namespace Data.DataImport.SqlRecordImporters
{
	/// <summary>
	/// Class EnumerableSqlRecord.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.DataImport.Interfaces.IEnumerableSqlRecord" />
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class EnumerableSqlRecord<T> : IEnumerableSqlRecord
	{
		/// <summary>
		/// The _row index
		/// </summary>
		protected int _rowIndex = 0;

		/// <summary>
		/// The _data
		/// </summary>
		protected IEnumerable<dynamic> _data;

		/// <summary>
		/// The _custom validator
		/// </summary>
		protected CustomValidation<T> _customValidator;

		/// <summary>
		/// The _custom SQL mapper
		/// </summary>
		protected CustomMapSqlRecord<T> _customSqlMapper;

		/// <summary>
		/// The _mapper
		/// </summary>
		protected FileValuesMap<T> _fileValuesMapper;
		/// <summary>
		/// The _import type
		/// </summary>
		protected Type _importType;
		/// <summary>
		/// The _SQL meta data
		/// </summary>
		protected List<SqlMetaData> _sqlMetadata;
		/// <summary>
		/// The _error column
		/// </summary>
		protected string _errorColumn;

		/// <summary>
		/// The _validator
		/// </summary>
		protected IObjectValidator _validator;
		/// <summary>
		/// Gets or sets the SQL metadata.
		/// </summary>
		/// <value>The SQL metadata.</value>
		public List<SqlMetaData> SqlMetadata
		{
			get { return _sqlMetadata; }
			internal set { _sqlMetadata = value; }
		}

		/// <summary>
		/// The _has rows
		/// </summary>
		protected bool _hasRows = false;
		/// <summary>
		/// Determines whether this instance has rows.
		/// </summary>
		/// <returns><c>true</c> if this instance has rows; otherwise, <c>false</c>.</returns>
		public virtual bool HasRows()
		{
			return _hasRows;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="EnumerableSqlRecord{T}"/> class.
		/// </summary>
		/// <param name="parameters">The parameters.</param>
		/// <param name="hasRows">if set to <c>true</c> [has rows].</param>
		public EnumerableSqlRecord(IImportEnumerableParameters<T> parameters, bool hasRows)
		{
			if (parameters == null) { throw new ArgumentNullException("parameters"); }
			this._importType = typeof(T);
			this._data = parameters.Data;
			this._hasRows = hasRows; 
			this.SqlMetadata = new List<SqlMetaData>();
			if (parameters.SqlMetadata != null) { this.SqlMetadata.AddRange(parameters.SqlMetadata); }
			this._errorColumn = parameters.ErrorColumnName;
			this._customSqlMapper = parameters.CustomSqlMapper;
			this._customValidator = parameters.CustomValidator;
			this._validator = parameters.Validator;

		}

		/// <summary>
		/// Initializes a new instance of the <see cref="EnumerableSqlRecord{T}"/> class.
		/// </summary>
		/// <param name="parameters">The parameters.</param>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0")]
		public EnumerableSqlRecord(IImportEnumerableParameters<T> parameters) : this(parameters, false)
		{
			this._hasRows = _data != null && _data.Any();
		}

		/// <summary>
		/// Returns an enumerator that iterates through the collection.
		/// </summary>
		/// <returns>A <see cref="T:System.Collections.Generic.IEnumerator`1" /> that can be used to iterate through the collection.</returns>
		public IEnumerator<SqlDataRecord> GetEnumerator()
		{
			if (_data == null || !_data.Any()) { yield break; }

			PropertyInfo[] properties = _importType.GetProperties();
			StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;
			this._validator = this._validator ?? new ObjectValidator();
			bool? isDynamicType = null;

			int errorColumnOrdinal = -1;
			var sqlMetaArray = _sqlMetadata.ToArray();

			if (_sqlMetadata.Any(x => comparer.Equals(x.Name, _errorColumn)))
			{
				SqlDataRecord tempRecord = new SqlDataRecord(sqlMetaArray);
				errorColumnOrdinal = tempRecord.GetOrdinal(_errorColumn); //will cause an exception if it does not exist, hence the any check
				tempRecord = null;
			}

			foreach (dynamic row in _data)
			{
				_rowIndex++;
				SqlDataRecord record = new SqlDataRecord(sqlMetaArray);
				List<string> errors = new List<string>();

				//check the first object to see if it is a dynamic type as we dont need to run it throught the object mapper in that case
				if (!isDynamicType.HasValue) { isDynamicType = FileIOHelpers.IsDynamicType(row); }

				T rowObj = default(T);

				if (isDynamicType.Value)
				{
					try
					{
						rowObj = FileIOUtilities.MapObject<T>(row, _rowIndex, _validator, _fileValuesMapper, ref errors);
					}
					catch (Exception ex)
					{
						errors.Add(ex.ToString());
					}
				}
				else
				{
					rowObj = row;
				}

				try
				{
					//built in data annotation validation
					this._validator.TryValidate(rowObj, ref errors);
					//custom validation
					if (_customValidator != null)
					{
						_customValidator.Invoke(rowObj, ref errors);
					}
				}
				catch (Exception ex)
				{
					errors.Add(ex.ToString());
				}

				ISqlRecordMapper mapperObj = null;
				//if they provide a custom mapper use that one over the interface.
				if (this._customSqlMapper != null)
				{
					this._customSqlMapper.Invoke(rowObj, record, _rowIndex, errors);
				}
				else if ((mapperObj = rowObj as ISqlRecordMapper) != null)
				{
					mapperObj.MapSqlRecord(record, _rowIndex, errors);
				}
				else //last ditch effort, hopefully they don't rely on this
				{
					object val;
					//try to set the rows from the metadata, and the properties
					foreach (SqlMetaData metaData in _sqlMetadata)
					{
						string name = metaData.Name;
						val = null;
						if (!comparer.Equals(name, _errorColumn))
						{
							var prop = properties.FirstOrDefault(x => comparer.Equals(x.Name, name));
							if (prop != null && (val = prop.GetValue(rowObj, null)) != null)
							{
								record.SetValue(record.GetOrdinal(name), val);
							}
						}
					}
					//if an error column is defined, set the import errors
					if (errorColumnOrdinal != -1 && errors.Count != 0)
					{
						string errorMessage = FileIOHelpers.ErrorsToXml(errors, _rowIndex);
						record.SetString(errorColumnOrdinal, errorMessage);
					}
				}
				yield return record;
			}
		}

		/// <summary>
		/// Returns an enumerator that iterates through a collection.
		/// </summary>
		/// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.GetEnumerator();
		}
	}
}
