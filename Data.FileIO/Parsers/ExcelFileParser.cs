// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ExcelFileParser.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Interfaces;
using Data.FileIO.Core;

namespace Data.FileIO.Parsers
{
    /// <summary>
    /// Class ExcelFileParser.
    /// </summary>
    /// <seealso cref="Data.FileIO.Interfaces.IExcelFileParser" />
    public class ExcelFileParser : IExcelFileParser
    {
        /// <summary>
        /// The _file name
        /// </summary>
        private string _fileName;

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
        /// Gets or sets the header row.
        /// </summary>
        /// <value>The header row.</value>
        public int HeaderRow { get; set; }

        /// <summary>
        /// Gets or sets the data row start.
        /// </summary>
        /// <value>The data row start.</value>
        public int DataRowStart { get; set; }

        /// <summary>
        /// Gets or sets the data row end.
        /// </summary>
        /// <value>The data row end.</value>
        public int DataRowEnd { get; set; } = -1;

        /// <summary>
        /// Gets or sets the name of the sheet.
        /// </summary>
        /// <value>The name of the sheet.</value>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets the row start.
        /// </summary>
        /// <value>The row start.</value>
        public int RowStart
        {
            get { return this.DataRowStart; }
        }

        /// <summary>
        /// The 1st column in the sheet that is valid for parsing, If null, then A will be assumed. Useful when a sheet contains multiple data sets.
        /// </summary>
        /// <value>The start column key.</value>
        public string StartColumnKey { get; set; }

        /// <summary>
        /// The very last column that should be parsed. If null, the entire sheet will be parsed. Useful when a sheet contains multiple data sets.
        /// </summary>
        /// <value>The end column key.</value>
        public string EndColumnKey { get; set; }

        /// <summary>
        /// Prevents a default instance of the <see cref="ExcelFileParser" /> class from being created.
        /// </summary>
        // ReSharper disable once UnusedMember.Local
        private ExcelFileParser() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelFileParser" /> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="headerRow">The header row.</param>
        /// <param name="dataRowStart">The data row start.</param>
        public ExcelFileParser(string fileName, string sheetName, int headerRow = 1, int dataRowStart = 2)
        {
            this.FileName = fileName;
            this.SheetName = sheetName;
            this.HeaderRow = headerRow;
            this.DataRowStart = dataRowStart;
        }

        /// <summary>
        /// Determines whether this instance has rows.
        /// </summary>
        /// <returns><c>true</c> if this instance has rows; otherwise, <c>false</c>.</returns>
        /// <exception cref="System.NotImplementedException"></exception>
        public bool HasRows()
        {
            var comparer = StringComparer.InvariantCultureIgnoreCase;

            using (var spreadsheetDocument = SpreadsheetDocument.Open(this.FileName, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                //find the sheet with the matching name
                var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(x => comparer.Equals(x.Name, this.SheetName));
                if (sheet == null) { return false; }

                //this is used by the reader to load the sheet for processing
                var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                //used to get the rowcount of the sheet
                // ReSharper disable once PossibleNullReferenceException
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                var hasDataRow = sheetData.Descendants<Row>().Any(row =>
                    row.RowIndex >= this.DataRowStart &&
                    row.Descendants<Cell>().Any(cell => cell.CellValue != null && !string.IsNullOrWhiteSpace(cell.CellValue.Text))
                );
                if (!hasDataRow)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Gets the headers for the specified worksheet.
        /// </summary>
        /// <param name="removeInvalidChars">if set to <c>true</c> [remove invalid chars].</param>
        /// <returns>IDictionary&lt;System.String, System.String&gt;.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
        public IDictionary<string, string> GetHeaders(bool removeInvalidChars = false)
        {
            var comparer = StringComparer.InvariantCultureIgnoreCase;
            // dictionary for the headers, the key will end up being the cell address minus the row number so that the headers and values can match up by key
            IDictionary<string, string> headers = new Dictionary<string, string>(comparer);


            Package spreadsheetPackage = null;

            var startColumnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(this.StartColumnKey);
            var endColumnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(this.EndColumnKey);

            try
            {
                spreadsheetPackage = Package.Open(_fileName, FileMode.Open, FileAccess.Read);

                using (var spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    //find the sheet with the matching name
                    var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(x => comparer.Equals(x.Name, this.SheetName));
                    if (sheet == null) { return headers; }

                    //this is used by the reader to load the sheet for processing
                    var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;

                    //needed to look up text values from cells
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sst = null;
                    if (sstpart != null) { sst = sstpart.SharedStringTable; }

                    //open the reader from the part
                    var reader = OpenXmlReader.Create(worksheetPart);
                    uint rowIndex = 0;


                    while (reader.Read())
                    {
                        //read until we find our rows, then loop through them
                        if (reader.ElementType != typeof(Row)) { continue; }
                        if (rowIndex > this.HeaderRow) { break; }

                        do
                        {
                            var row = (Row)reader.LoadCurrentElement();
                            rowIndex = row.RowIndex;

                            if (rowIndex > this.HeaderRow) { break; }
                            if (rowIndex != this.HeaderRow) { continue; }

                            if (row.HasChildren)
                            {
                                foreach (var cell in row.Descendants<Cell>())
                                {
                                    var cellKey = FileIOUtilities.GetColumnKey(cell.CellReference.Value);
                                    if (startColumnIndex != -1 || endColumnIndex != -1)
                                    {
                                        int cellIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(cellKey);
                                        if (startColumnIndex >= 0 && cellIndex < startColumnIndex)
                                        {
                                            continue;
                                        }
                                        if (endColumnIndex >= 0 && cellIndex > endColumnIndex)
                                        {
                                            break;
                                        }
                                    }

                                    string value;

                                    if (cell.CellValue != null)
                                    {
                                        if (cell.DataType != null && cell.DataType == CellValues.SharedString && sst != null)
                                        {
                                            //read the text value out of the shared string table.
                                            value = sst.ChildElements[int.Parse(cell.CellValue.Text)].InnerText;
                                        }
                                        else
                                        {
                                            value = cell.CellValue.Text;
                                        }

                                        if (rowIndex == this.HeaderRow)
                                        {
                                            headers.Add(cellKey, value);
                                        }
                                    }
                                }
                            }

                            if (rowIndex != this.HeaderRow) continue;

                            if (removeInvalidChars)
                            {
                                //remove all characters that are not allowed in .net property names
                                headers = FileIOUtilities.FixupHeaders(headers);
                            }
                        } while (reader.ReadNextSibling());
                        //rows are all done, break out of the loop
                        break;
                    }
                }

            }
            finally
            {
                if(spreadsheetPackage!= null)
                {
                    spreadsheetPackage.Close();
                }
            }


            return headers;
        }

        /// <summary>
        /// Gets the sheetnames for the work book.
        /// </summary>
        /// <param name="fileName">The excel file to read the sheet names from.</param>
        /// <param name="includeHidden">if set to <c>true</c> [returns hidden sheets].</param>
        /// <returns>IList&lt;System.String&gt;.</returns>
        public static IList<string> GetSheetNames(string fileName, bool includeHidden = false)
        {
            IList<string> ret = new List<string>();
            Package spreadsheetPackage = null;
            try
            {
                spreadsheetPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read);

                using (var spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage))
                {
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where((item) => includeHidden || (!includeHidden && !item.IsHidden()));
                    foreach(var sheet in sheets)
                    {
                        ret.Add(sheet.Name);
                    }
                }
            }
            finally
            {
                if (spreadsheetPackage != null)
                {
                    spreadsheetPackage.Close();
                }
            }

            return ret;
        }
        /// <summary>
        /// Parses the file.
        /// </summary>
        /// <returns>IEnumerable&lt;dynamic&gt;.</returns>
        /// <exception cref="System.IO.FileNotFoundException"></exception>
        public IEnumerable<dynamic> ParseFile()
        {
            var comparer = StringComparer.InvariantCultureIgnoreCase;

            Package spreadsheetPackage = null;

            var startColumnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(this.StartColumnKey);
            var endColumnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(this.EndColumnKey);

            try
            {
                spreadsheetPackage = Package.Open(_fileName, FileMode.Open, FileAccess.Read);

                using (var spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    //find the sheet with the matching name
                    var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(x => comparer.Equals(x.Name, this.SheetName));
                    if (sheet == null) { yield break; }

                    //this is used by the reader to load the sheet for processing
                    var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                    //used to get the rowcount of the sheet
                    // ReSharper disable once PossibleNullReferenceException
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    //check to ensure that we have any data rows at all, that have cell values
                    var hasDataRow = sheetData.Descendants<Row>().Any(row =>
                        row.RowIndex >= this.DataRowStart &&
                        row.Descendants<Cell>().Any(cell => cell.CellValue != null && !string.IsNullOrWhiteSpace(cell.CellValue.Text))
                    );
                    if (!hasDataRow)
                    {
                        yield break;
                    }

                    //needed to look up text values from cells
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sst = null;
                    if (sstpart != null)
                    {
                        sst = sstpart.SharedStringTable;
                    }

                    var cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                    IList<NumberingFormat> numberingFormats = null;
                    if (workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null)
                    {
                        numberingFormats = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.OfType<NumberingFormat>().ToList();
                    }

                    //open the reader from the part
                    var reader = OpenXmlReader.Create(worksheetPart);

                    // dictionary for the headers, the key will end up being the cell address minus the row number so that the headers and values can match up by key
                    IDictionary<string, string> headers = new Dictionary<string, string>();
                    //the values dictionary for each row
                    Dictionary<string, string> values = null;

                    while (reader.Read())
                    {
                        //read until we find our rows, then loop through them
                        if (reader.ElementType == typeof(Row))
                        {
                            do
                            {
                                var row = (Row)reader.LoadCurrentElement();
                                uint rowIndex = row.RowIndex;

                                if (!(rowIndex == this.HeaderRow) && rowIndex < this.DataRowStart) { continue; }
                                //if they have specified a end read row bail out if the rowindex exceeds that end value
                                if (this.DataRowEnd >= this.DataRowStart && rowIndex > this.DataRowEnd) { break; }

                                if (row.HasChildren)
                                {
                                    //loop through all of the cells in the row, building a list of either header keys, or value keys depending on which row it is.
                                    values = new Dictionary<string, string>();
                                    foreach (var cell in row.Descendants<Cell>())
                                    {
                                        var cellKey = FileIOUtilities.GetColumnKey(cell.CellReference.Value);

                                        if (startColumnIndex != -1 || endColumnIndex != -1)
                                        {
                                            var cellIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(cellKey);
                                            if (startColumnIndex >= 0 && cellIndex < startColumnIndex)
                                            {
                                                continue;
                                            }
                                            if (endColumnIndex >= 0 && cellIndex > endColumnIndex)
                                            {
                                                break;
                                            }
                                        }

                                        var value = String.Empty;

                                        if (cell.DataType != null && cell.DataType == CellValues.SharedString && sst != null)
                                        {
                                            //read the text value out of the shared string table.
                                            value = sst.ChildElements[int.Parse(cell.CellValue.Text)].InnerText;
                                        }
                                        else if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.Text))
                                        {
                                            if (cell.StyleIndex != null) //style index?
                                            {
                                                //frakking excel dates. wth. determing if a cell is formatted as a date is a huge pita
                                                var cellFormat = (CellFormat)cellFormats.ElementAt((int)cell.StyleIndex.Value);
                                                NumberingFormat numberingFormat = null;
                                                if (numberingFormats != null && cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.HasValue)
                                                {
                                                    numberingFormat =
                                                        numberingFormats.FirstOrDefault(fmt => fmt.NumberFormatId.Value == cellFormat.NumberFormatId.Value);
                                                }

                                                if ((cell.DataType != null && cell.DataType == CellValues.Date) //just in case
                                                    ||
                                                    (cellFormat.NumberFormatId != null &&
                                                     (cellFormat.NumberFormatId >= 14 && cellFormat.NumberFormatId <= 22)) //built in date formats
                                                    ||
                                                    (numberingFormat != null &&
                                                    !numberingFormat.FormatCode.Value.Contains("[") && //so we dont match [Red] in numbering formats.... /sigh
                                                     Regex.IsMatch(numberingFormat.FormatCode, "d|h|m|s|y", RegexOptions.IgnoreCase))
                                                    //custom date formats, would an isdate be too hard msft?
                                                    ) // Dates
                                                {
                                                    value = Convert.ToString(DateTime.FromOADate(double.Parse(cell.CellValue.Text)),
                                                        CultureInfo.InvariantCulture);
                                                }
                                                else
                                                {
                                                    value = cell.CellValue.Text;
                                                }
                                            }
                                            else
                                            {
                                                value = cell.CellValue.Text;
                                            }
                                        }

                                        if (rowIndex >= this.DataRowStart)
                                        {
                                            values.Add(cellKey, (value ?? "").Trim());
                                        }
                                        else if (rowIndex == this.HeaderRow)
                                        {
                                            headers.Add(cellKey, value);
                                        }
                                    }
                                }

                                //we have accumulated either our headers or values for this row, so now we need to handle them
                                if (rowIndex >= this.DataRowStart)
                                {
                                    //sometimes excel reports the last row as higher than it actually has values, and we end up with an empty row. 
                                    //skip the row if this happens. otherwise we can output a weird object value
                                    if (values.Any(x => !String.IsNullOrWhiteSpace(x.Value)))
                                    {
                                        dynamic retObj = FileIOUtilities.RowToExpando(headers, values, Convert.ToInt32(rowIndex));
                                        retObj.RowId = Convert.ToInt32(rowIndex);
                                        //stream the data row back to the caller
                                        yield return retObj;
                                    }
                                }
                                else if (rowIndex == this.HeaderRow)
                                {
                                    //remove all characters that are not allowed in .net property names
                                    headers = FileIOUtilities.FixupHeaders(headers);
                                    //string headersString = "\t[" + String.Join("] varchar(500),\r\n\t[", headers.Values) + "] varchar(500)";
                                    //Debug.WriteLine(headersString);
                                }
                            } while (reader.ReadNextSibling());
                            //rows are all done, break out of the loop
                            break;
                        }
                    }
                }

            }
            finally
            {
                spreadsheetPackage.Close();
            }
        }
    }

}
