// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ExpandoDataReader.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Data.DataImport.Readers
{
	/// <summary>
	/// Class ExpandoDataReaderExtensions.
	/// </summary>
	public static class ExpandoDataReaderExtensions
	{
		/// <summary>
		/// Gets the expando data reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="list">The list.</param>
		/// <returns>ExpandoDataReader&lt;T&gt;.</returns>
		public static ExpandoDataReader<T> GetExpandoDataReader<T>(this IEnumerable<T> list)
		{
			return new ExpandoDataReader<T>(list);
		}
	}

	/// <summary>
	/// Will wrap a an enumerable&lt;dynamic&gt; within a IDataReader.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="System.Data.IDataReader" />
	/// <seealso cref="System.Data.IDataRecord" />
	/// <seealso cref="System.IDisposable" />
	public class ExpandoDataReader<T> : IDataReader, IDataRecord, IDisposable
	{
		/// <summary>
		/// The _enumerator
		/// </summary>
		private IEnumerator<T> _enumerator = null;
		//used by the getordinal function to provide a fast ordinal lookup
		/// <summary>
		/// The _name lookup
		/// </summary>
		private Dictionary<string, int> _nameLookup = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
		/// <summary>
		/// The _first
		/// </summary>
		IDictionary<string, object> _first;
		/// <summary>
		/// The _depth
		/// </summary>
		private int _depth = -1;
		/// <summary>
		/// The _closed
		/// </summary>
		private bool _closed = false;

		/// <summary>
		/// Initializes a new instance of the <see cref="ExpandoDataReader{T}" /> class.
		/// </summary>
		/// <param name="list">The list.</param>
		public ExpandoDataReader(IEnumerable<T> list)
		{
			if (list != null)
			{
				this._enumerator = list.GetEnumerator();
				if (_enumerator.MoveNext())
				{
					//certain operations could use the data first without performing a read, so we need to peek at the data without advancing the enumerator
					_first = (IDictionary<string, object>)_enumerator.Current;
					//_enumerator.Reset(); //throws exception so we will handle using the depth

					for (int i = 0; i < _first.Keys.Count; i++)
					{
						_nameLookup[_first.Keys.ElementAt(i)] = i;
					}
				}
			}
		}

		/// <summary>
		/// Gets the current row.
		/// </summary>
		/// <value>The current row.</value>
		private IDictionary<string, object> CurrentRow
		{
			get
			{
				//if we have not moved forward, return the first record as the current
				if (_depth <= 0)
				{
					return _first;
				}

				return _enumerator.Current as IDictionary<string, object>;
			}
		}

		#region IDataReader Members

		/// <summary>
		/// Closes the <see cref="T:System.Data.IDataReader" /> Object.
		/// </summary>
		public void Close()
		{
			if (!_closed)
			{
				_nameLookup.Clear();
				_closed = true;
				_enumerator.Dispose();
			}
		}

		/// <summary>
		/// Gets a value indicating the depth of nesting for the current row.
		/// </summary>
		/// <value>The depth.</value>
		public int Depth
		{
			get { return _depth; }
		}

		/// <summary>
		/// Returns a <see cref="T:System.Data.DataTable" /> that describes the column metadata of the <see cref="T:System.Data.IDataReader" />.
		/// </summary>
		/// <returns>A <see cref="T:System.Data.DataTable" /> that describes the column metadata.</returns>
		public DataTable GetSchemaTable()
		{
			DataTable ret = new DataTable();
			Type colType = typeof(object);
			foreach (var key in _first.Keys)
			{
				ret.Columns.Add(new DataColumn(key, colType));
			}
			return ret;
		}

		/// <summary>
		/// Gets a value indicating whether the data reader is closed.
		/// </summary>
		/// <value><c>true</c> if this instance is closed; otherwise, <c>false</c>.</value>
		public bool IsClosed
		{
			get { return _closed; }
		}

		/// <summary>
		/// Advances the data reader to the next result, when reading the results of batch SQL statements.
		/// </summary>
		/// <returns>true if there are more rows; otherwise, false.</returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public bool NextResult()
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Advances the <see cref="T:System.Data.IDataReader" /> to the next record.
		/// </summary>
		/// <returns>true if there are more rows; otherwise, false.</returns>
		public bool Read()
		{
			//don't move next on the very first record as we already did a move next
			if (_depth++ == -1)
			{
				return true;
			}
			return _enumerator.MoveNext();
		}

		/// <summary>
		/// Gets the number of rows changed, inserted, or deleted by execution of the SQL statement.
		/// </summary>
		/// <value>The records affected.</value>
		public int RecordsAffected
		{
			get { return -1; }
		}

		#endregion

		#region IDisposable Members

		/// <summary>
		/// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1816:CallGCSuppressFinalizeCorrectly")]
		public void Dispose()
		{
			Close();
			GC.SuppressFinalize(true); 
		}
		/// <summary>
		/// Releases unmanaged and - optionally - managed resources.
		/// </summary>
		/// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				Close();
			}
		}
		#endregion

		#region IDataRecord Members

		/// <summary>
		/// Gets the number of columns in the current row.
		/// </summary>
		/// <value>The field count.</value>
		public int FieldCount
		{
			get
			{
				return _first.Keys.Count;
			}
		}

		/// <summary>
		/// Gets the value of the specified column as a Boolean.
		/// </summary>
		/// <param name="i">The zero-based column ordinal.</param>
		/// <returns>The value of the column.</returns>
		public bool GetBoolean(int i)
		{
			return (bool)GetValue(i);
		}

		/// <summary>
		/// Gets the 8-bit unsigned integer value of the specified column.
		/// </summary>
		/// <param name="i">The zero-based column ordinal.</param>
		/// <returns>The 8-bit unsigned integer value of the specified column.</returns>
		public byte GetByte(int i)
		{
			return (byte)GetValue(i);
		}

		/// <summary>
		/// Reads a stream of bytes from the specified column offset into the buffer as an array, starting at the given buffer offset.
		/// </summary>
		/// <param name="i">The zero-based column ordinal.</param>
		/// <param name="fieldOffset">The index within the field from which to start the read operation.</param>
		/// <param name="buffer">The buffer into which to read the stream of bytes.</param>
		/// <param name="bufferoffset">The index for <paramref name="buffer" /> to start the read operation.</param>
		/// <param name="length">The number of bytes to read.</param>
		/// <returns>The actual number of bytes read.</returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Gets the character value of the specified column.
		/// </summary>
		/// <param name="i">The zero-based column ordinal.</param>
		/// <returns>The character value of the specified column.</returns>
		public char GetChar(int i)
		{
			return (char)GetValue(i);
		}

		/// <summary>
		/// Reads a stream of characters from the specified column offset into the buffer as an array, starting at the given buffer offset.
		/// </summary>
		/// <param name="i">The zero-based column ordinal.</param>
		/// <param name="fieldoffset">The index within the row from which to start the read operation.</param>
		/// <param name="buffer">The buffer into which to read the stream of bytes.</param>
		/// <param name="bufferoffset">The index for <paramref name="buffer" /> to start the read operation.</param>
		/// <param name="length">The number of bytes to read.</param>
		/// <returns>The actual number of characters read.</returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Returns an <see cref="T:System.Data.IDataReader" /> for the specified column ordinal.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The <see cref="T:System.Data.IDataReader" /> for the specified column ordinal.</returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public IDataReader GetData(int i)
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Gets the data type information for the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The data type information for the specified field.</returns>
		/// <exception cref="System.NotImplementedException"></exception>
		public string GetDataTypeName(int i)
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Gets the date and time data value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The date and time data value of the specified field.</returns>
		public DateTime GetDateTime(int i)
		{
			return (DateTime)GetValue(i);
		}

		/// <summary>
		/// Gets the fixed-position numeric value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The fixed-position numeric value of the specified field.</returns>
		public decimal GetDecimal(int i)
		{
			return (decimal)GetValue(i);
		}

		/// <summary>
		/// Gets the double-precision floating point number of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The double-precision floating point number of the specified field.</returns>
		public double GetDouble(int i)
		{
			return (double)GetValue(i);
		}

		/// <summary>
		/// Gets the <see cref="T:System.Type" /> information corresponding to the type of <see cref="T:System.Object" /> that would be returned from <see cref="M:System.Data.IDataRecord.GetValue(System.Int32)" />.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The <see cref="T:System.Type" /> information corresponding to the type of <see cref="T:System.Object" /> that would be returned from <see cref="M:System.Data.IDataRecord.GetValue(System.Int32)" />.</returns>
		public Type GetFieldType(int i)
		{
			return CurrentRow.ElementAt(i).Value.GetType();
		}

		/// <summary>
		/// Gets the single-precision floating point number of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The single-precision floating point number of the specified field.</returns>
		public float GetFloat(int i)
		{
			return (float)GetValue(i);
		}

		/// <summary>
		/// Returns the GUID value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The GUID value of the specified field.</returns>
		public Guid GetGuid(int i)
		{
			return (Guid)GetValue(i);
		}

		/// <summary>
		/// Gets the 16-bit signed integer value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The 16-bit signed integer value of the specified field.</returns>
		public short GetInt16(int i)
		{
			return (short)GetValue(i);
		}

		/// <summary>
		/// Gets the 32-bit signed integer value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The 32-bit signed integer value of the specified field.</returns>
		public int GetInt32(int i)
		{
			return (int)GetValue(i);
		}

		/// <summary>
		/// Gets the 64-bit signed integer value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The 64-bit signed integer value of the specified field.</returns>
		public long GetInt64(int i)
		{
			return (long)GetValue(i);
		}

		/// <summary>
		/// Gets the name for the field to find.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The name of the field or the empty string (""), if there is no value to return.</returns>
		public string GetName(int i)
		{
			return CurrentRow.ElementAt(i).Key;
		}

		/// <summary>
		/// Return the index of the named field.
		/// </summary>
		/// <param name="name">The name of the field to find.</param>
		/// <returns>The index of the named field.</returns>
		public int GetOrdinal(string name)
		{
			if (_nameLookup.ContainsKey(name))
			{
				return _nameLookup[name];
			}
			else
			{
				return -1;
			}
		}

		/// <summary>
		/// Gets the string value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The string value of the specified field.</returns>
		public string GetString(int i)
		{
			return (string)GetValue(i);
		}

		/// <summary>
		/// Return the value of the specified field.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>The <see cref="T:System.Object" /> which will contain the field value upon return.</returns>
		public object GetValue(int i)
		{
			return CurrentRow.ElementAt(i).Value;
		}

		/// <summary>
		/// Populates an array of objects with the column values of the current record.
		/// </summary>
		/// <param name="values">An array of <see cref="T:System.Object" /> to copy the attribute fields into.</param>
		/// <returns>The number of instances of <see cref="T:System.Object" /> in the array.</returns>
		public int GetValues(object[] values)
		{
			int getValues = 0;
			if (values != null)
			{
				getValues = Math.Max(FieldCount, values.Length);

				for (int i = 0; i < getValues; i++)
				{
					values[i] = GetValue(i);
				}
			}
			return getValues;
		}

		/// <summary>
		/// Return whether the specified field is set to null.
		/// </summary>
		/// <param name="i">The index of the field to find.</param>
		/// <returns>true if the specified field is set to null; otherwise, false.</returns>
		public bool IsDBNull(int i)
		{
			return GetValue(i) == null;
		}

		/// <summary>
		/// Gets the column with the specified name.
		/// </summary>
		/// <param name="name">The name.</param>
		/// <returns>System.Object.</returns>
		public object this[string name]
		{
			get
			{
				return GetValue(GetOrdinal(name));
			}
		}

		/// <summary>
		/// Gets the column located at the specified index.
		/// </summary>
		/// <param name="i">The i.</param>
		/// <returns>System.Object.</returns>
		public object this[int i]
		{
			get
			{
				return GetValue(i);
			}
		}

		#endregion
	}

}
