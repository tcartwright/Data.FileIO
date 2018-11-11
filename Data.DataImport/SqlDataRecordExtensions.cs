// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 05-07-2016
//
// Last Modified By : tdcart
// Last Modified On : 05-07-2016
// ***********************************************************************
// <copyright file="SqlDataRecordExtensions.cs" company="Tim Cartwright">
//     Copyright © Tim Cartwright 2016
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.Server;

namespace Data.DataImport
{
	/// <summary>
	/// Class SqlDataRecordExtensions.
	/// </summary>
	public static class SqlDataRecordExtensions
	{
		/// <summary>
		/// Sets the boolean.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">if set to <c>true</c> [value].</param>
		public static void SetBoolean(this SqlDataRecord record, string fieldName, bool? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetBoolean(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the byte.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetByte(this SqlDataRecord record, string fieldName, byte? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetByte(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the bytes.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="fieldOffset">The field offset.</param>
		/// <param name="buffer">The buffer.</param>
		/// <param name="bufferOffset">The buffer offset.</param>
		/// <param name="length">The length.</param>
		public static void SetBytes(this SqlDataRecord record, string fieldName, long fieldOffset, byte[] buffer, int bufferOffset, int length)
		{
			int ordinal = GetOrdinal(record, fieldName);
			record.SetBytes(ordinal, fieldOffset, buffer, bufferOffset, length);
		}
		/// <summary>
		/// Sets the character.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetChar(this SqlDataRecord record, string fieldName, char? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetChar(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the chars.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="fieldOffset">The field offset.</param>
		/// <param name="buffer">The buffer.</param>
		/// <param name="bufferOffset">The buffer offset.</param>
		/// <param name="length">The length.</param>
		public static void SetChars(this SqlDataRecord record, string fieldName, long fieldOffset, char[] buffer, int bufferOffset, int length)
		{
			int ordinal = GetOrdinal(record, fieldName);
			record.SetChars(ordinal, fieldOffset, buffer, bufferOffset, length);
		}
		/// <summary>
		/// Sets the date time.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetDateTime(this SqlDataRecord record, string fieldName, DateTime? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				//only set date if in valid range: https://msdn.microsoft.com/en-us/library/ms187819.aspx
				// January 1, 1753, through December 31, 9999
				if (value.Value > new DateTime(1753, 1, 1) && value.Value < new DateTime(9999, 12, 31))
				{
					record.SetDateTime(ordinal, value.Value);
				}
				else
				{
					throw new ArgumentOutOfRangeException("value", "The date provided is out of range for Sql Server. The date must be between January 1, 1753, through December 31, 9999.");
				}
			}
		}
		/// <summary>
		/// Sets the date time offset.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetDateTimeOffset(this SqlDataRecord record, string fieldName, DateTimeOffset? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetDateTimeOffset(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the database null.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		public static void SetDBNull(this SqlDataRecord record, string fieldName)
		{
			int ordinal = GetOrdinal(record, fieldName);
			record.SetDBNull(ordinal);
		}
		/// <summary>
		/// Sets the decimal.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetDecimal(this SqlDataRecord record, string fieldName, decimal? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetDecimal(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the double.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetDouble(this SqlDataRecord record, string fieldName, double? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetDouble(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the float.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetFloat(this SqlDataRecord record, string fieldName, float? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetFloat(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the unique identifier.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetGuid(this SqlDataRecord record, string fieldName, Guid? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetGuid(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the int16.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetInt16(this SqlDataRecord record, string fieldName, Int16? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetInt16(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the int32.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetInt32(this SqlDataRecord record, string fieldName, Int32? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetInt32(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the int64.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetInt64(this SqlDataRecord record, string fieldName, Int64? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetInt64(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the string.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetString(this SqlDataRecord record, string fieldName, string value)
		{
			int ordinal = GetOrdinal(record, fieldName);

			if (value != null)
			{
				record.SetString(ordinal, value);
			}
		}
		/// <summary>
		/// Sets the time span.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetTimeSpan(this SqlDataRecord record, string fieldName, TimeSpan? value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			if (value.HasValue)
			{
				record.SetTimeSpan(ordinal, value.Value);
			}
		}
		/// <summary>
		/// Sets the value.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <param name="value">The value.</param>
		public static void SetValue<T>(this SqlDataRecord record, string fieldName, T value)
		{
			int ordinal = GetOrdinal(record, fieldName);
			record.SetValue(ordinal, value);
		}

		/// <summary>
		/// Gets the ordinal.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="fieldName">Name of the field.</param>
		/// <returns>System.Int32.</returns>
		/// <exception cref="System.ArgumentNullException">record
		/// or
		/// fieldName</exception>
		private static int GetOrdinal(SqlDataRecord record, string fieldName)
		{
			if (record == null) throw new ArgumentNullException("record");
			if (String.IsNullOrWhiteSpace(fieldName)) throw new ArgumentNullException("fieldName");
			int ordinal = record.GetOrdinal(fieldName);
			return ordinal;
		}
	}
}
