// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="CsvWriter.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Data.FileIO.Interfaces;
using Data.FileIO.Core;
using System.Data;
using Data.FileIO.Common.Utilities;

namespace Data.FileIO.Writers
{
	/// <summary>
	/// Class CsvWriter.
	/// </summary>
	public class CsvWriter
	{
		/// <summary>
		/// The first header format
		/// </summary>
		private const string FirstHeaderFormat	= "{0}";
		/// <summary>
		/// The second header format
		/// </summary>
		private const string SecondHeaderFormat = ",{0}";
		/// <summary>
		/// The first item format
		/// </summary>
		private const string FirstItemFormat	= "{0}{1}{0}"; //the zero placeholders are the quoted ids for the item formats
		/// <summary>
		/// The second item format
		/// </summary>
		private const string SecondItemFormat	= ",{0}{1}{0}";

		/// <summary>
		/// Writes the data to a csv file.
		/// </summary>
		/// <param name="destinationPath">The destination path.</param>
		/// <param name="data">The data.</param>
		/// <param name="quotedFieldCharacter">The quoted field character. Will apply to all fields in the file.</param>
		/// <param name="writeHeaders">if set to <c>true</c> [write headers].</param>
		/// <exception cref="System.ArgumentNullException">data</exception>
		public virtual void WriteData(string destinationPath, IDataReader data, char? quotedFieldCharacter = null, bool writeHeaders = true)
		{
			if (data == null) { throw new ArgumentNullException("data"); }

			Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
			string quote = quotedFieldCharacter.HasValue ? quotedFieldCharacter.Value.ToString() : String.Empty;
			bool isFirstItem = true;
			string format = null;
			List<string> fields = new List<string>();

			for (int i = 0; i < data.FieldCount; i++)
			{
				fields.Add(data.GetName(i));
			}

			using (StreamWriter stream = File.CreateText(destinationPath))
			{
				if (writeHeaders)
				{
					foreach (var field in fields)
					{
						format = isFirstItem ? FirstHeaderFormat : SecondHeaderFormat;
						stream.Write(String.Format(format, field));
						isFirstItem = false;
					}
					stream.WriteLine();
				}

				while (data.Read())
				{
					isFirstItem = true;

					foreach (var field in fields)
					{
						int ordinal = data.GetOrdinal(field);
						object val = null;
						if (!data.IsDBNull(ordinal)) { val = data.GetValue(ordinal); }

						format = isFirstItem ? FirstItemFormat : SecondItemFormat;
						stream.Write(String.Format(format, quote, val));
						isFirstItem = false;
					}
					stream.WriteLine();
				}
			}
		}


		/// <summary>
		/// Writes the data to a csv file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="destinationPath">The destination path.</param>
		/// <param name="data">The data.</param>
		/// <param name="columnInfoList">The column information list.</param>
		/// <param name="writeHeaders">if set to <c>true</c> [write headers].</param>
		/// <exception cref="System.ArgumentNullException">
		/// data
		/// or
		/// columnInfoList
		/// </exception>
		public virtual void WriteData<T>(string destinationPath, IDataReader data, CsvColumnInfoList<T> columnInfoList, bool writeHeaders = true)
		{
			if (data == null) { throw new ArgumentNullException("data"); }
			var dataList = FileIOHelpers.GetEnumerableFromReader<T>(data);
			this.WriteData<T>(destinationPath, dataList, columnInfoList, writeHeaders);
		}

		/// <summary>
		/// Writes the data to a csv file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="destinationPath">The destination path.</param>
		/// <param name="data">The data.</param>
		/// <param name="columnInfoList">The column information list.</param>
		/// <param name="writeHeaders">if set to <c>true</c> [write headers].</param>
		/// <exception cref="System.ArgumentNullException">
		/// data
		/// or
		/// columnInfoList
		/// </exception>
		public virtual void WriteData<T>(string destinationPath, IEnumerable<T> data, CsvColumnInfoList<T> columnInfoList, bool writeHeaders = true)
		{
			if (data == null) { throw new ArgumentNullException("data"); }
			if (columnInfoList == null) { throw new ArgumentNullException("columnInfoList"); }

			Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
			bool isFirstItem = true;
			string format = null;

			using (StreamWriter stream = File.CreateText(destinationPath))
			{
				if (writeHeaders)
				{
					foreach (var colInfo in columnInfoList)
					{
						format = isFirstItem ? FirstHeaderFormat : SecondHeaderFormat;
						stream.Write(String.Format(format, colInfo.HeaderName));
						isFirstItem = false;
					}
					stream.WriteLine();
				}

				foreach (var item in data)
				{
					isFirstItem = true;
					foreach (var colInfo in columnInfoList)
					{
						format = isFirstItem ? FirstItemFormat : SecondItemFormat;
						stream.Write(String.Format(format, colInfo.QuotedFieldCharacter, colInfo.ValueFunction.Invoke(item)));
						isFirstItem = false;
					}
					stream.WriteLine();
				}
			}
		}

		/// <summary>
		/// Writes the data to a csv file. Uses reflection to dump out all the properties of the object to the file.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="destinationPath">The destination path.</param>
		/// <param name="data">The data.</param>
		/// <param name="quotedFieldCharacter">The quoted field character. Will apply to all fields in the file.</param>
		/// <param name="writeHeaders">if set to <c>true</c> [write headers].</param>
		/// <exception cref="System.ArgumentNullException">data</exception>
		public virtual void WriteData<T>(string destinationPath, IEnumerable<T> data, char? quotedFieldCharacter = null, bool writeHeaders = true)
		{
			if (data == null) { throw new ArgumentNullException("data"); }

			Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
			var type = typeof(T);
			var properties = type.GetProperties().Where(pi => pi.GetGetMethod() != null);
			string quote = quotedFieldCharacter.HasValue ? quotedFieldCharacter.Value.ToString() : String.Empty;
			bool isFirstItem = true;
			string format = null;

			using (StreamWriter stream = File.CreateText(destinationPath))
			{
				if (writeHeaders)
				{
					foreach (var propertyInfo in properties)
					{
						format = isFirstItem ? FirstHeaderFormat : SecondHeaderFormat;
						stream.Write(String.Format(format, propertyInfo.Name));
						isFirstItem = false;
					}
					stream.WriteLine();
				}

				foreach (var item in data)
				{
					isFirstItem = true;
					foreach (var propertyInfo in properties)
					{
						format = isFirstItem ? FirstItemFormat : SecondItemFormat;
						stream.Write(String.Format(format, quote, propertyInfo.GetValue(item, null)));
						isFirstItem = false;
					}
					stream.WriteLine();
				}
			}
		}
	}
}
