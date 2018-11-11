// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="CsvFileParser.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Data.FileIO.Common;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Core;
using Data.FileIO.Interfaces;
using Microsoft.VisualBasic.FileIO;

namespace Data.FileIO.Parsers
{
	/// <summary>
	/// Class CSVFileParser.
	/// </summary>
	/// <seealso cref="Data.FileIO.Interfaces.ICsvFileParser" />
	public class CsvFileParser : ICsvFileParser
	{
		/// <summary>
		/// The _file name
		/// </summary>
		private string _fileName;
		/// <summary>
		/// The _first row contains headers
		/// </summary>
		private bool _firstRowContainsHeaders;
		/// <summary>
		/// The _rows to skip
		/// </summary>
		private int _rowsToSkip;
		/// <summary>
		/// The _has quoted fields
		/// </summary>
		private bool _hasQuotedFields;
		/// <summary>
		/// The _trim white space
		/// </summary>
		private bool _trimWhiteSpace;
		/// <summary>
		/// The _fixed widths
		/// </summary>
		private int[] _fixedWidths;
		/// <summary>
		/// The _delimiters
		/// </summary>
		private string[] _delimiters;

		/// <summary>
		/// Gets or sets the fixed widths.
		/// </summary>
		/// <value>The fixed widths.</value>
		public int[] FixedWidths
		{
			get { return _fixedWidths; }
			set { _fixedWidths = value; }
		}

		/// <summary>
		/// Gets or sets the delimiters.
		/// </summary>
		/// <value>The delimiters.</value>
		public string[] Delimiters
		{
			get { return _delimiters; }
			set { _delimiters = value; }
		}

		/// <summary>
		/// Gets or sets the name of the file.
		/// </summary>
		/// <value>The name of the file.</value>
		/// <exception cref="System.IO.FileNotFoundException"></exception>
		public string FileName
		{
			get { return _fileName; }
			set 
			{
				if (!File.Exists(value)) { throw new FileNotFoundException(value); }

				_fileName = value; 
			}
		}

		/// <summary>
		/// Gets or sets a value indicating whether [first row contains headers].
		/// </summary>
		/// <value><c>true</c> if [first row contains headers]; otherwise, <c>false</c>.</value>
		public bool FirstRowContainsHeaders
		{
			get { return _firstRowContainsHeaders; }
			set { _firstRowContainsHeaders = value; }
		}

		/// <summary>
		/// Gets or sets the rows to skip.
		/// </summary>
		/// <value>The rows to skip.</value>
		public int RowsToSkip
		{
			get { return _rowsToSkip; }
			set { _rowsToSkip = value; }
		}

		/// <summary>
		/// Gets or sets a value indicating whether this instance has quoted fields.
		/// </summary>
		/// <value><c>true</c> if this instance has quoted fields; otherwise, <c>false</c>.</value>
		public bool HasQuotedFields
		{
			get { return _hasQuotedFields; }
			set { _hasQuotedFields = value; }
		}

		/// <summary>
		/// Gets or sets a value indicating whether [trim white space].
		/// </summary>
		/// <value><c>true</c> if [trim white space]; otherwise, <c>false</c>.</value>
		public bool TrimWhiteSpace
		{
			get { return _trimWhiteSpace; }
			set { _trimWhiteSpace = value; }
		}

		/// <summary>
		/// Prevents a default instance of the <see cref="CsvFileParser" /> class from being created.
		/// </summary>
		private CsvFileParser() { }

		/// <summary>
		/// Gets the row start.
		/// </summary>
		/// <value>The row start.</value>
		public int RowStart
		{
			get
			{
				int rowstart = this.FirstRowContainsHeaders ? 2 : 1;
				rowstart += this.RowsToSkip;
				return rowstart;
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CsvFileParser" /> class.
		/// </summary>
		/// <param name="fileName">Name of the file.</param>
		/// <param name="firstRowContainsHeaders">if set to <c>true</c> [first row contains headers].</param>
		/// <param name="rowsToSkip">The rows to skip.</param>
		/// <param name="hasQuotedFields">if set to <c>true</c> [has quoted fields].</param>
		/// <param name="trimWhiteSpace">if set to <c>true</c> [trim white space].</param>
		public CsvFileParser(string fileName, bool firstRowContainsHeaders = true, int rowsToSkip = 0, bool hasQuotedFields = true, bool trimWhiteSpace = true)
		{
			this.FileName = fileName;
			this.FirstRowContainsHeaders = firstRowContainsHeaders;
			this.RowsToSkip = rowsToSkip;
			this.HasQuotedFields = hasQuotedFields;
			this.TrimWhiteSpace = trimWhiteSpace;
		}

		/// <summary>
		/// Determines whether this instance has rows.
		/// </summary>
		/// <returns><c>true</c> if this instance has rows; otherwise, <c>false</c>.</returns>
		public bool HasRows()
		{
			int lineCount = 0;
			int dataRowStart = this.RowStart;

			using (var reader = File.OpenText(this.FileName))
			{
				while (reader.ReadLine() != null)
				{
					if (++lineCount >= dataRowStart)
					{
						return true;
					}
				}
			}

			return false;
		}

		/// <summary>
		/// Parses the file.
		/// </summary>
		/// <returns>IEnumerable&lt;dynamic&gt;.</returns>
		/// <exception cref="System.IO.FileNotFoundException"></exception>
		public IEnumerable<dynamic> ParseFile()
		{
			int rowIndex = this.RowStart;

			// TextFieldParser is in the Microsoft.VisualBasic.FileIO namespace.
			using (TextFieldParser parser = new TextFieldParser(_fileName))
			{
				parser.SetFieldWidths(_fixedWidths.Maybe().FirstOrDefault());
				parser.SetDelimiters(_delimiters.Maybe().DefaultIfEmpty(new string[] { ",", "\t" }).First());
				parser.TextFieldType = parser.FieldWidths == null ? FieldType.Delimited : FieldType.FixedWidth;

				parser.TrimWhiteSpace = TrimWhiteSpace;
				parser.HasFieldsEnclosedInQuotes = HasQuotedFields;
				parser.CommentTokens = new string[] { "#" };

				string[] headers = null;
				string[] values = null;

				if (FirstRowContainsHeaders) { headers = FileIOUtilities.FixupHeaders(parser.ReadFields()); }

				for (int i = 0; i < RowsToSkip; i++)
				{
					parser.ReadLine();
				}

				while (!parser.EndOfData)
				{
					values = parser.ReadFields();

					if (headers == null)
					{
						headers = Enumerable.Range(0, values.Length).Select(x => "Col" + x).ToArray<string>();
					}
					//if all of the values are empty in the row, do not emit it. just continue on
					if (values.Any(x => !String.IsNullOrWhiteSpace(x)))
					{
						dynamic retObj = FileIOUtilities.RowToExpando(headers, values);
						retObj.RowId = rowIndex++;
						yield return retObj;
					}
				}

				parser.Close();
			}
		}

	}

}
