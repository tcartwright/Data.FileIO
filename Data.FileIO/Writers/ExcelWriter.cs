// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ExcelWriter.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Data.FileIO.Common;
using Data.FileIO.Common.Utilities;
using Data.FileIO.Core;
using Data.FileIO.Interfaces;
using DL = Data.FileIO.Core.DynamicLinq;
// ReSharper disable ArrangeThisQualifier
// ReSharper disable VirtualMemberCallInConstructor
// ReSharper disable EmptyStatement


namespace Data.FileIO.Writers
{
    /// <summary>
    /// Class ExcelWriter.
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]
    public class ExcelWriter : IDisposable
    {
        #region fields
        /// <summary>
        /// Class FormatsKey.
        /// </summary>
        private class FormatsKey
        {
            /// <summary>
            /// The hash code
            /// </summary>
            private readonly Guid _hashCode;
            /// <summary>
            /// Gets or sets the type of the cell.
            /// </summary>
            /// <value>The type of the cell.</value>
            public ExcelCellType CellType { get; set; }
            /// <summary>
            /// Gets or sets the format code.
            /// </summary>
            /// <value>The format code.</value>
            public string FormatCode { get; set; }
            /// <summary>
            /// Gets or sets a value indicating whether this instance is default.
            /// </summary>
            /// <value><c>true</c> if this instance is default; otherwise, <c>false</c>.</value>
            public bool IsDefault { get; set; }
            /// <summary>
            /// Gets or sets the number format identifier.
            /// </summary>
            /// <value>The number format identifier.</value>
            public int NumberFormatId { get; set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="FormatsKey"/> class.
            /// </summary>
            public FormatsKey()
            {
                this._hashCode = Guid.NewGuid();
            }
            /// <summary>
            /// Determines whether the specified <see cref="System.Object" /> is equal to this instance.
            /// </summary>
            /// <param name="obj">The object to compare with the current object.</param>
            /// <returns><c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.</returns>
            public override bool Equals(object obj)
            {
                // If parameter is null return false.
                if (obj == null)
                {
                    return false;
                }

                // If parameter cannot be cast return false.
                var p = obj as FormatsKey;
                if (p == null)
                {
                    return false;
                }

                // Return true if the fields match:
                return (this.CellType == p.CellType)
                    && (this.FormatCode == p.FormatCode)
                    && (this.IsDefault == p.IsDefault)
                    && (this.NumberFormatId == p.NumberFormatId);
            }
            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table.</returns>
            public override int GetHashCode()
            {
                return this._hashCode.GetHashCode();
            }
            /// <summary>
            /// Returns a <see cref="System.String" /> that represents this instance.
            /// </summary>
            /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
            public override string ToString()
            {
                return string.Format("{0}-{1}", this.CellType, this.FormatCode);
            }
        }
        /// <summary>
        /// The comparer
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
        // ReSharper disable once InconsistentNaming
        protected readonly StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;
        /// <summary>
        /// The _spread sheet
        /// </summary>
        private SpreadsheetDocument _spreadSheet;
        /// <summary>
        /// The _spreadsheet package
        /// </summary>
        private Package _spreadsheetPackage;
        /// <summary>
        /// The _custom formats
        /// </summary>
        private Dictionary<FormatsKey, uint> _customFormats = new Dictionary<FormatsKey, uint>();
        /// <summary>
        /// The _template path
        /// </summary>
        private readonly string _templatePath;
        /// <summary>
        /// The _header row
        /// </summary>
        private int _headerRow;
        /// <summary>
        /// The _data row
        /// </summary>
        private int _dataRow;
        /// <summary>
        /// Will hold the last rows written for each sheet
        /// </summary>
        private Dictionary<string, int> _lastRowsWritten = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
        /// <summary>
        /// The _generate headers from type
        /// </summary>
        private bool _generateHeadersFromType;
        /// <summary>
        /// The _create sheet if not found
        /// </summary>
        private bool _createSheetIfNotFound;
        /// <summary>
        /// The _is closed
        /// </summary>
        private bool _isClosed;
        /// <summary>
        /// The _default cell width
        /// </summary>
        private double _defaultCellWidth = 9.7109375; //yanked out of an existing sheet
        #endregion fields

        #region Properties
        /// <summary>
        /// The memory stream that will hold the spreadsheet document as it is written to.
        /// </summary>
        /// <value>The spreadsheet stream.</value>
        public MemoryStream SpreadsheetStream { get; private set; } // The stream that the spreadsheet gets returned on

        /// <summary>
        /// Gets the template path.
        /// </summary>
        /// <value>The template path.</value>
        public string TemplatePath
        {
            get { return _templatePath; }
        }

        /// <summary>
        /// Gets or sets the header row.
        /// </summary>
        /// <value>The header row.</value>
        public int HeaderRow
        {
            get { return _headerRow; }
            set { _headerRow = value; }
        }

        /// <summary>
        /// Gets or sets the data row.
        /// </summary>
        /// <value>The data row.</value>
        public int DataRow
        {
            get { return _dataRow; }
            set { _dataRow = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [generate headers from type]. Has no effect unless the header row is less than 1.
        /// </summary>
        /// <value><c>true</c> if [generate headers from type]; otherwise, <c>false</c>.</value>
        public bool GenerateHeadersFromType
        {
            get { return _generateHeadersFromType; }
            set { _generateHeadersFromType = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [create sheet if not found].
        /// </summary>
        /// <value><c>true</c> if [create sheet if not found]; otherwise, <c>false</c>.</value>
        public bool CreateSheetIfNotFound
        {
            get { return _createSheetIfNotFound; }
            set { _createSheetIfNotFound = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [clear sheet data before write].
        /// </summary>
        /// <value><c>true</c> if [clear sheet data before write]; otherwise, <c>false</c>.</value>
        public bool ClearSheetDataBeforeWrite { get; set; }
        #endregion Properties

        #region ctors
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter" /> class. Will create or update a sheet based upon a template.
        /// </summary>
        /// <param name="templatePath">The template source path.</param>
        /// <param name="headerRow">The header row. If the header row is zero or less all properties from the object will be output to the sheet.</param>
        /// <param name="dataRow">The data row.</param>
        /// <exception cref="System.ArgumentNullException">templateSourcePath
        /// or
        /// destinationPath</exception>
        /// <exception cref="System.IO.FileNotFoundException">templateSourcePath</exception>
        /// <exception cref="System.ArgumentException">The data row must be greater than the header row.</exception>
        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "templatePath"), SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ExcelWriter(string templatePath, int headerRow = 1, int dataRow = 2)
        {
            if (string.IsNullOrWhiteSpace(templatePath)) { throw new ArgumentNullException("templatePath"); }
            if (!File.Exists(templatePath)) { throw new FileNotFoundException("templatePath"); }
            if (dataRow < headerRow) { throw new ArgumentException("The data row must be greater than the header row."); }

            _templatePath = templatePath;
            _dataRow = dataRow;
            _headerRow = headerRow;
            this.GenerateHeadersFromType = _headerRow < 1;
            this.ClearSheetDataBeforeWrite = true;

            var templateBytes = File.ReadAllBytes(_templatePath);
            this.SpreadsheetStream = new MemoryStream();
            this.SpreadsheetStream.Write(templateBytes, 0, templateBytes.Length);
            _spreadsheetPackage = Package.Open(this.SpreadsheetStream, FileMode.Open, FileAccess.ReadWrite);
            _spreadSheet = SpreadsheetDocument.Open(_spreadsheetPackage);
            AddWorkbookStyles(_spreadSheet.WorkbookPart);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter" /> class. Creates a new excel file without using a base template.
        /// </summary>
        /// <param name="headerRow">The header row.</param>
        /// <param name="dataRow">The data row.</param>
        /// <exception cref="System.ArgumentException">The data row must be greater than the header row.</exception>
        [SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ExcelWriter(int headerRow = 1, int dataRow = 2)
        {
            if (dataRow < headerRow) { throw new ArgumentException("The data row must be greater than the header row."); }

            _dataRow = dataRow;
            _headerRow = headerRow;
            this.GenerateHeadersFromType = headerRow < 1;
            this.CreateSheetIfNotFound = true; //set this by default as the entire workbook is new

            this.SpreadsheetStream = new MemoryStream();
            _spreadSheet = SpreadsheetDocument.Create(this.SpreadsheetStream, SpreadsheetDocumentType.Workbook);
            //need to add the workbook part as it wont be there for a new sheet
            var workBookPart = _spreadSheet.AddWorkbookPart();
            workBookPart.Workbook = new Workbook();
            //workBookPart.Workbook.Append(new FileVersion() { ApplicationName = "ExcelWriter" });
            //add the sheets collection
            workBookPart.Workbook.AppendChild(new Sheets());
            AddWorkbookStyles(workBookPart);

            _spreadSheet.Package.Flush();
            _spreadSheet.WorkbookPart.Workbook.Save();
        }

        #endregion ctors

        #region WriteDataToSheet methods
        /// <summary>
        /// Writes the data to sheet using the <see cref="IDataReader" /> as a source.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="reader">The reader.</param>
        /// <param name="columnInfoList">A list of optional column informational objects used to transform the output of the sheet. The value property must match a lambda expression format.</param>
        /// <param name="appendData">if set to <c>true</c> [append data] to the existing sheet.</param>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "2")]
        public virtual void WriteDataToSheet(string sheetName, IDataReader reader, IList<IColumnInfoBase> columnInfoList = null, bool appendData = false)
        {
            var stringType = typeof(string);
            //first convert the reader to a type
            var readerType = FileIOHelpers.GetTypeFromReader(reader);

            //now we are going to reflect into the write sheet method that takes a type and a reader
            // ReSharper disable once RedundantTypeArgumentsOfMethod
            var writeToSheetMethod = new Action<string, IDataReader, ColumnInfoList<object>, bool>(this.WriteDataToSheet<object>);
            var writeToSheetGenericMethod = writeToSheetMethod.Method.GetGenericMethodDefinition().MakeGenericMethod(readerType);

            //create a List<ColumnHeaderMap<T>> dynamically
            var columnHeaderMapType = typeof(ColumnInfo<>).MakeGenericType(readerType);
            var columnHeaderMapTypeCtor = columnHeaderMapType.GetConstructors().First(c =>
            {
                var parms = c.GetParameters();
                return parms.Count() == 3
                    && parms[0].ParameterType == stringType
                    && parms[1].ParameterType == stringType
                    && parms[2].ParameterType.BaseType == typeof(LambdaExpression);
            });

            var mapsType = typeof(ColumnInfoList<>).MakeGenericType(readerType);
            //create a type of the new dictionary to pass in to the other method that we can add our compiled maps to
            var maps = Activator.CreateInstance(mapsType, true);

            //take the string headers map, and convert them on the fly to lambdas and a IList<IColumnHeaderMap<object>>
            if (!columnInfoList.IsEmpty())
            {
                //create a generic version of the lambda parser method
                var parseLambdaMethod = new Func<string, ParameterExpression, Expression<Func<object, object>>>(this.ParseLambda<object, object>);
                var parseLambdaGenericMethod = parseLambdaMethod.Method.GetGenericMethodDefinition().MakeGenericMethod(readerType, stringType);

                //get the add method off the list
                // ReSharper disable once PossibleNullReferenceException
                var mapsAddMethod = maps.GetType().BaseType.GetMethod("Add");

                var readerParam = Expression.Parameter(readerType, "obj");
                var baseFuncType = typeof(Func<,>).MakeGenericType(readerType, typeof(object));

                // ReSharper disable once PossibleNullReferenceException
                foreach (var map in columnInfoList)
                {
                    //create the generic lambda body from the function string
                    var lambda = ((LambdaExpression)parseLambdaGenericMethod.Invoke(this, new object[] { map.ValueFunctionString, readerParam }));
                    //build the Expression<Func<readerType, object>> dynamically
                    var expr = Expression.Lambda(baseFuncType, lambda.Body, readerParam);

                    var mapInstance = columnHeaderMapTypeCtor.Invoke(new object[] { map.ColumnName, map.HeaderName, expr }) as IColumnInfoBase;
                    // ReSharper disable once PossibleNullReferenceException
                    mapInstance.CellType = map.CellType;
                    mapInstance.FormatCode = map.FormatCode;
                    mapInstance.UpdateHeader = map.UpdateHeader;

                    //add the lambda in
                    mapsAddMethod.Invoke(maps, new object[] { mapInstance });
                }
            }

            //invoke it passing our dynamic reader type, and the other parameters along
            writeToSheetGenericMethod.Invoke(this, new[] { sheetName, reader, maps, appendData });
        }

        /// <summary>
        /// Writes the data to a sheet using the <see cref="IDataReader" /> as a source with the supplied type.
        /// </summary>
        /// <typeparam name="T">If the type passed in extends the <see cref="Data.FileIO.Common.Interfaces.IDataRecordMapper" /> it will be used to map the <see cref="IDataReader" /> fields to the type.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="reader">The reader.</param>
        /// <param name="columnInfoList">A list of optional column informational objects used to transform the output of the sheet.</param>
        /// <param name="appendData">if set to <c>true</c> [append data] to the existing sheet.</param>
        public virtual void WriteDataToSheet<T>(string sheetName, IDataReader reader, ColumnInfoList<T> columnInfoList = null, bool appendData = false)
        {
            //get the type to convert the reader to
            var readerType = typeof(T);

            //get the method of the other WriteDataToSheet method
            // ReSharper disable once RedundantTypeArgumentsOfMethod
            var writeToSheetMethod = new Action<string, IEnumerable<object>, ColumnInfoList<object>, bool>(this.WriteDataToSheet<object>);
            var writeToSheetGenericMethod = writeToSheetMethod.Method.GetGenericMethodDefinition().MakeGenericMethod(readerType);

            //get the method that will allow us to convert a reader to an IEnumerable<T>
            var readerToEnumerableMethod = new Func<IDataReader, IEnumerable<object>>(FileIOHelpers.GetEnumerableFromReader<object>);
            var readerToEnumerableGenericMethod = readerToEnumerableMethod.Method.GetGenericMethodDefinition().MakeGenericMethod(readerType);

            //call the convert reader to enumerable to pass in to the other sheet
            var readerData = readerToEnumerableGenericMethod.Invoke(null, new object[] { reader });
            //invoke the call FINALLY
            writeToSheetGenericMethod.Invoke(this, new[] { sheetName, readerData, columnInfoList, appendData });
        }

        /// <summary>
        /// Writes the data to an Excel sheet.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="data">The data.</param>
        /// <param name="columnInfoList">A list of optional column informational objects used to transform the output of the sheet.</param>
        /// <param name="appendData">if set to <c>true</c> [append data] to the existing sheet.</param>
        /// <exception cref="ObjectDisposedException">The writer is closed.</exception>
        /// <exception cref="ArgumentException">The DataRow must be greater than the HeaderRow.</exception>
        /// <exception cref="ArgumentNullException">sheetName
        /// or
        /// data</exception>
        /// <exception cref="FileFormatException">Worksheet named ' + sheetName + ' not found in the workbook.</exception>
        /// <exception cref="FormatException">The writer is closed.</exception>
        /// <exception cref="FileFormatException">The DataRow must be greater than the HeaderRow.</exception>
        [SuppressMessage("Microsoft.Performance", "CA1809:AvoidExcessiveLocals")]
        [SuppressMessage("Microsoft.Maintainability", "CA1505:AvoidUnmaintainableCode"),
            SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "DocumentFormat.OpenXml.Spreadsheet.CellValue.#ctor(System.String)"),
            SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling"),
            SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        public virtual void WriteDataToSheet<T>(string sheetName, IEnumerable<T> data, ColumnInfoList<T> columnInfoList = null, bool appendData = false)
        {
            if (_isClosed) { throw new ObjectDisposedException("The writer is closed."); }
            if (this.DataRow < this.HeaderRow) { throw new ArgumentException("The DataRow must be greater than the HeaderRow."); }
            if (string.IsNullOrWhiteSpace(sheetName)) { throw new ArgumentNullException("sheetName"); }
            if (data == null) { throw new ArgumentNullException("data"); }

            var createdSheet = false;
            //if they are appending to an existing sheet, get the last row index, else use the data row
            var rowIndex = appendData && _lastRowsWritten.ContainsKey(sheetName) ? _lastRowsWritten[sheetName] : this.DataRow;
            //maintain the current data row to be used when writing in case we are appending
            var currentDataRow = rowIndex;

            var workbookPart = _spreadSheet.WorkbookPart;

            // - added deletion of calculation part becuase if it exists leaving it in and deleting rows can cause sheet corruption
            if (workbookPart.CalculationChainPart != null)
            {
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }

            #region get existing sheet or create a new sheet
            //find the sheet with the matching name
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(x => comparer.Equals(x.Name, sheetName));
            WorksheetPart worksheetPart;

            if (sheet != null)
            {
                worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
            }
            else
            {
                if (_createSheetIfNotFound)
                {
                    worksheetPart = CreateNewSheet(workbookPart, sheetName);
                    createdSheet = true;
                }
                else
                {
                    throw new FileFormatException("Worksheet named '" + sheetName + "' not found in the workbook.");
                }
            }
            #endregion get existing sheet or create a new sheet

            //get the sheetdata element so we can add rows to it below
            // ReSharper disable once PossibleNullReferenceException
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            var columns = worksheetPart.Worksheet.Descendants<Columns>().FirstOrDefault();
            if (columns == null)
            {
                //need to insert the columns BEFORE the sheet data else the sheet gets corrupted
                columns = worksheetPart.Worksheet.InsertBefore(new Columns(), sheetData);
                //make sure to add a default column else the sheet will be corrupted if no columns get added
                // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                columns.Append(new Column
                {
                    Min = 1U,
                    Max = 1U,
                    Width = _defaultCellWidth
                });
            }

            //we are going to cache the column styles so we don't have to retrieve the style from the column each row, 
            //Tuple: item1 == the style index, item2 == the cell type
            var styleIndexes = new Dictionary<string, Tuple<uint, CellValues>>(comparer);
            //list of all the headers where a numbering format has been added. To keep multiple numbering format adds from occurring
            var headerNumberFormatsAdded = new List<string>();

            //remove any existing rows that are in the data rows as leaving them in may corrupt the sheet
            if (this.ClearSheetDataBeforeWrite)
            {
                foreach (var row in sheetData.Descendants<Row>().Where(x => x.RowIndex >= currentDataRow))
                {
                    row.Remove();
                }
            }

            #region get the column headers and possibly update them
            //get the properties off the type as we may use them to write the headers
            var properties = typeof(T).GetProperties().Where(pi => pi.GetGetMethod() != null);

            IDictionary<string, string> headers;
            var propertyInfos = properties as PropertyInfo[] ?? properties.ToArray();
            if (_headerRow > 0 && !createdSheet)
            {
                //get the headers from the sheet itself if we did not create it.
                headers = GetHeaders(workbookPart, sheetData, _headerRow);
            }
            else
            {
                headers = this.GenerateHeadersFromType
                    ? GetHeadersFromType(propertyInfos) //auto generate the headers off the properties
                    : new Dictionary<string, string>(comparer); //they chose not to auto gen headers, so just roll with it
            }

            //if the headers ARE still empty and the column info maps are empty, we are going to auto gen the headers no matter what as otherwise we end up with an empty sheet.
            if (headers.IsEmpty() && columnInfoList.IsEmpty())
            {
                headers = GetHeadersFromType(propertyInfos);
            }

            //add or update existing headers to the header row in the sheet
            if (this.HeaderRow > 0)
            {
                // ReSharper disable once RedundantTypeArgumentsOfMethod
                AddOrUpdateHeaders<T>(columnInfoList, createdSheet, sheetData, headers);
            }
            #endregion get the column headers and possibly update them

            //add any custom formats in the maps to the style sheet before looping all the data
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (!columnInfoList.IsEmpty())
            {
                // ReSharper disable once AssignNullToNotNullAttribute
                foreach (var map in columnInfoList.Where(m => m.CellType == ExcelCellType.Quoted || (m.CellType != ExcelCellType.General && !string.IsNullOrWhiteSpace(m.FormatCode))))
                {
                    if (map.CellType != ExcelCellType.Quoted)
                    {
                        AddNumberFormat(stylesPart, map.CellType, map.FormatCode);
                    }
                    else
                    {
                        AddCellFormat(stylesPart, ExcelCellType.Quoted, false);
                    }
                }
            }

            //loop through the data adding it to the sheet
            foreach (var item in data)
            {
                var addRow = false;
                //create a new row
                var row = sheetData.Descendants<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);

                if (row == null)
                {
                    addRow = true;
                    row = new Row();
                }
                //loop through the excel file column headers looking for data to write to the column
                foreach (var header in headers)
                {
                    var columnName = header.Key; //this MUST end up being an excel column name
                    var headerName = header.Value;

                    string propertyValue = null;
                    var propertyType = typeof(string);
                    uint styleIndex = 0;

                    //this excelCellType, formatCode, and mapValueMethod will be supplied by the column info if we have a column info. Default to General and empty
                    var excelCellType = ExcelCellType.General;
                    var formatCode = string.Empty;
                    Func<T, object> mapValueMethod = null;

                    #region Get the value and type of cell
                    if (!columnInfoList.IsEmpty())
                    {
                        //search for a mapping function in the custom maps first 
                        var map = columnInfoList.FirstOrDefault(m => m.IsMatch(columnName, headerName));
                        if (map != null)
                        {
                            mapValueMethod = map.ValueFunction;
                            excelCellType = map.CellType;
                            formatCode = map.FormatCode;
                            propertyType = map.ValueFunctionType;
                        }
                    }

                    if (mapValueMethod != null)
                    {
                        var propValue = mapValueMethod.Invoke(item);

                        if (propValue != null) //if we get null keep on trucking
                        {
                            propertyValue = Convert.ToString(propValue);
                            //if we are on the first row, and the cell type is general see if we can match the property type to a cell type
                            if (excelCellType == ExcelCellType.General && propertyType != typeof(string))
                            {
                                //try to interpret the ExcelCellType from the propertyType
                                excelCellType = GetExcelCellTypeFromPropertyType(propertyType);
                                //only add the format if we have a format code and we are processing the first data row
                                if (!headerNumberFormatsAdded.Contains(headerName.ToLower()) && excelCellType != ExcelCellType.General && !string.IsNullOrWhiteSpace(formatCode))
                                {
                                    AddNumberFormat(stylesPart, excelCellType, formatCode);
                                    headerNumberFormatsAdded.Add(headerName.ToLower());
                                }
                            }
                        }
                    }
                    else
                    {
                        //find the property that matches up to the the header
                        var property = propertyInfos.FirstOrDefault(pi => comparer.Equals(pi.Name, headerName));

                        //property not found, so continue on with the next header
                        if (property == null) { continue; }

                        //get the value from the property and object
                        propertyValue = Convert.ToString(property.GetValue(item, null));
                        var notNullType = Nullable.GetUnderlyingType(property.PropertyType);
                        propertyType = notNullType ?? property.PropertyType;
                    }
                    #endregion Get the value and type of cell

                    ////no value so continue on with the next header
                    //if (string.IsNullOrWhiteSpace(propertyValue)) { continue; }
                    //convert the property type to an excel cellvalues enum value
                    var cellType = GetCellTypeFromType(propertyType);

                    #region get the column style
                    //try to get the column style from the cache first, should be empty only for the first row
                    if (styleIndexes.ContainsKey(columnName))
                    {
                        styleIndex = styleIndexes[columnName].Item1;
                        cellType = styleIndexes[columnName].Item2;
                    }
                    else
                    {
                        //this should only be occurring on or after the first data row as it should be cached after the first time.
                        if (rowIndex >= currentDataRow)
                        {
                            //try to determine the cell type and format index from the type info
                            DetermineStyleIndex(worksheetPart, columnName, propertyType, excelCellType, formatCode, ref styleIndex, ref cellType);
                            //add it to the cache, even if zero for general style so we don't have to do this again on the next row.
                            styleIndexes.Add(columnName, new Tuple<uint, CellValues>(styleIndex, cellType));
                        }
                    }
                    #endregion get the column style

                    #region special date value formatting
                    //if we have a date we need to properly format it so bakka excel will display it as a date
                    //this conversion needs to occur before figuring out the style as the cell type will change to Number 
                    //then as formating of a date has to be done as a number. go figure.
                    if (propertyType == typeof(DateTime) || cellType == CellValues.Date || excelCellType == ExcelCellType.Date)
                    {
                        if (!string.IsNullOrWhiteSpace(propertyValue))
                        {
                            DateTime dateTemp;
                            if (DateTime.TryParse(propertyValue, out dateTemp))
                            {
                                //convert the date to excel format
                                propertyValue = dateTemp.ToOADate().ToString(CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                //not entirely sure what to do here. doesn't seem right to throw an exception and blow up the entire workbook/sheet write process
                                propertyValue = "#VALUE! : " + propertyValue;
                            }
                        }
                        //dates HAVE to be cell type number, else the sheet gets corrupted
                        cellType = CellValues.Number;
                    }
                    #endregion special date value formatting

                    var cell = row.Descendants<Cell>().FirstOrDefault(c => comparer.Equals(c.CellReference, columnName + rowIndex));

                    if (cell == null)
                    {
                        cell = new Cell()
                        {
                            DataType = cellType,
                            CellValue = new CellValue(propertyValue),
                            CellReference = columnName + rowIndex
                        };

                        //cells HAVE to be inserted in order, else guess what? yea. the sheet will get corrupted.
                        var siblingColumnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(columnName) + 1;
                        var siblingCell = row.Descendants<Cell>().FirstOrDefault(c => FileIOUtilities.ConvertExcelColumnNameToNumber(c.CellReference) >= siblingColumnIndex);

                        if (siblingCell == null)
                        {
                            row.AppendChild(cell);
                        }
                        else
                        {
                            row.InsertBefore(cell, siblingCell);
                        }

                        if (styleIndex != 0) { cell.StyleIndex = styleIndex; }
                    }
                    else
                    {
                        //cell.DataType = cellType;
                        cell.CellValue = new CellValue(propertyValue);
                    }
                }

                if (addRow)
                {
                    //rows HAVE to be inserted in order, else guess what? yea. the sheet will get corrupted.
                    var siblingRow = sheetData.Descendants<Row>().FirstOrDefault(r => r.RowIndex > rowIndex);
                    //write the row to the sheet data if it is not already appended
                    row.RowIndex = (uint)rowIndex;
                    if (siblingRow == null)
                    {
                        sheetData.AppendChild(row);
                    }
                    else
                    {
                        sheetData.InsertBefore(row, siblingRow);
                    }
                }

                rowIndex++;
            }


            //need to save everything when using memorystream, else we will lose the data
            worksheetPart.Worksheet.Save();
            _spreadSheet.WorkbookPart.Workbook.Save();
            //save the row index in case they want to append to the sheet
            _lastRowsWritten.AddOrUpdate(sheetName, rowIndex);
        }

        #endregion WriteDataToSheet methods

        #region write to methods
        /// <summary>
        /// Writes the work sheet to a file. The writer will be closed before performing the write operation.
        /// </summary>
        /// <param name="destinationPath">The destination path.</param>
        /// <exception cref="System.ArgumentNullException">destinationPath</exception>
        /// <exception cref="System.NullReferenceException">SpreadsheetStream</exception>
        /// <exception cref="System.NotSupportedException">The writer must be closed prior to writing to a file.</exception>
        [SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes"), SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "SpreadsheetStream")]
        public void WriteTo(string destinationPath)
        {
            if (string.IsNullOrWhiteSpace(destinationPath)) { throw new ArgumentNullException("destinationPath"); }
            if (this.SpreadsheetStream == null) { throw new NullReferenceException("SpreadsheetStream"); };

            this.Close();
            if (!_isClosed) { throw new NotSupportedException("The writer must be closed prior to writing to a destination."); }

            var destinationDirectory = Path.GetDirectoryName(destinationPath);
            //create the directory if it does not exist
            // ReSharper disable once AssignNullToNotNullAttribute
            Directory.CreateDirectory(destinationDirectory);

            this.SpreadsheetStream.Position = 0;
            using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
            {
                this.SpreadsheetStream.WriteTo(fileStream);
            }
        }

        /// <summary>
        /// Writes the spreadsheet to a stream. The writer will be closed before performing the write operation.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="System.NullReferenceException">SpreadsheetStream
        /// or
        /// stream</exception>
        /// <exception cref="System.NotSupportedException">The writer must be closed prior to writing.</exception>
        [SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes"), SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "SpreadsheetStream")]
        public void WriteTo(Stream stream)
        {
            if (this.SpreadsheetStream == null) { throw new NullReferenceException("SpreadsheetStream"); };
            if (stream == null) { throw new NullReferenceException("stream"); };

            this.Close();
            if (!_isClosed) { throw new NotSupportedException("The writer must be closed prior to writing to a destination."); }

            this.SpreadsheetStream.Position = 0;
            this.SpreadsheetStream.WriteTo(stream);
        }
        #endregion write to methods

        #region helpers
        //public IEnumerable<ValidationErrorInfo> ValidateSheet()
        //{
        //	var validator = new OpenXmlValidator();
        //	IEnumerable<ValidationErrorInfo> errors = validator.Validate(_spreadSheet);
        //	return errors;
        //}


        /// <summary>
        /// Determines the index of the style and adds it to the column if appropriate.
        /// </summary>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyType">Type of the property.</param>
        /// <param name="excelCellType">Type of the excel cell.</param>
        /// <param name="formatCode">The format code.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <param name="cellType">Type of the cell.</param>
        private void DetermineStyleIndex(WorksheetPart worksheetPart, string columnName, Type propertyType, ExcelCellType excelCellType, string formatCode, ref uint styleIndex, ref CellValues cellType)
        {
            var columns = worksheetPart.Worksheet.Descendants<Columns>().FirstOrDefault();
            var columnIndex = FileIOUtilities.ConvertExcelColumnNameToNumber(columnName);
            //as we have the ability to add headers, there may be indexes outside of the columns collection, or we may have no columns at all if creating a sheet
            // ReSharper disable once PossibleNullReferenceException
            var col = columns.Descendants<Column>().FirstOrDefault(c => columnIndex >= c.Min && columnIndex <= c.Max);
            if (col != null && col.Style != null) { styleIndex = col.Style; }

            //column style will be used if available, else try to figure out an applicable style from the data types
            if (styleIndex == 0)
            {
                if (excelCellType == ExcelCellType.Quoted)
                {
                    //this column will be prefixed with a ' and set to ignore numbers when stored as text
                    var error = new IgnoredError() { NumberStoredAsText = BooleanValue.FromBoolean(true) };
                    AddIgnoredError(worksheetPart, columnName, error);
                    cellType = CellValues.String;
                    styleIndex = GetFormatIndex(excelCellType, formatCode);
                }
                else if (cellType == CellValues.Date || excelCellType == ExcelCellType.Date)
                {
                    cellType = CellValues.Number;
                    styleIndex = GetFormatIndex(ExcelCellType.Date, formatCode);
                }
                //this only handles when the column is matched off a property on the object, else the != General will handle all other numerics
                else if (cellType == CellValues.Number)
                {
                    //select the decimal or numeric style
                    switch (Type.GetTypeCode(propertyType))
                    {
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            styleIndex = GetFormatIndex(ExcelCellType.Decimal, formatCode);
                            break;
                        default:
                            styleIndex = GetFormatIndex(ExcelCellType.Number, formatCode);
                            break;
                    }
                }
                else if (excelCellType != ExcelCellType.General) //at this point it should be one of the numeric formats or a general column
                {
                    cellType = CellValues.Number;
                    styleIndex = GetFormatIndex(excelCellType, formatCode);
                }

                //we found a style index so lets try to assign it to the column in question so even for new and empty cells the style will work
                //only do this code if the rowindex is the very first data row
                if (styleIndex > 0)
                {
                    if (col == null)
                    {
                        var colMin = columns.Descendants<Column>().Select(x => x.Max.Value).Max() + 1U;
                        var colMax = (uint)columnIndex;
                        //add the stretch column if needed. columns actually cover a range of columns in the sheet via min and max as a range
                        if (colMin < colMax)
                        {
                            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                            columns.Append(new Column
                            {
                                Min = colMin,
                                Max = colMax - 1U,
                                Width = _defaultCellWidth
                            });
                        }
                        // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                        columns.Append(new Column
                        {
                            Min = colMax,
                            Max = colMax,
                            Width = _defaultCellWidth,
                            Style = styleIndex
                        });
                    }
                    else
                    {
                        col.Style = styleIndex;
                    }
                }
            }
        }


        /// <summary>
        /// Adds the ignored error.
        /// </summary>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="error">The error.</param>
        private static void AddIgnoredError(WorksheetPart worksheetPart, string columnName, IgnoredError error)
        {
            var ignoredErrors = worksheetPart.Worksheet.Descendants<IgnoredErrors>().FirstOrDefault();
            if (ignoredErrors == null)
            {
                ignoredErrors = new IgnoredErrors();
                worksheetPart.Worksheet.AppendChild(ignoredErrors);
            }
            error.SequenceOfReferences = new ListValue<StringValue>();
            error.SequenceOfReferences.Items.Add(StringValue.FromString(string.Format("{0}1:{0}65535", columnName)));
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            ignoredErrors.Append(error);
        }

        /// <summary>
        /// For a new sheet adds the headers from the properties. The column infos are used to update or add headers as needed.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnInfoList">The column information list.</param>
        /// <param name="createdSheet">if set to <c>true</c> [created sheet].</param>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="headers">The headers.</param>
        private void AddOrUpdateHeaders<T>(ColumnInfoList<T> columnInfoList, bool createdSheet, SheetData sheetData, IDictionary<string, string> headers)
        {
            var isNewHeaderRow = false;
            var row = sheetData.Descendants<Row>().FirstOrDefault(r => r.RowIndex == this.HeaderRow);
            if (row == null)
            {
                isNewHeaderRow = true;
                row = new Row();
            }

            //add the headers if we created the sheet to the sheet data
            if (createdSheet)
            {
                var colindex = 1;
                foreach (var header in headers)
                {
                    var headerName = header.Value;

                    //add the new cell
                    var cell = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(headerName),
                        CellReference = FileIOUtilities.ConvertExcelNumberToLetter(colindex++) + this.HeaderRow
                    };
                    row.AppendChild(cell);
                }
            }

            //add or update the header values from the custom maps that have column names
            if (!columnInfoList.IsEmpty())
            {
                //gen a random string so that it should not match any existing column names or properties.
                var gennedValue = Guid.NewGuid().ToString();
                uint? styleIndex = null;

                //loop through all the custom headers that have valid excel column name (hopefully as any 3 characters will seem valid)
                foreach (var map in columnInfoList.Where(m => m.IsValidColumnName))
                {
                    //add or override the headers with our custom ones if we have any. 
                    //this will allow the consumer to add new columns to a sheet that already has headers
                    headers.AddOrUpdate(map.ColumnName, gennedValue);

                    var cellAddress = map.ColumnName + this.HeaderRow;
                    //search for a cell with this particular address
                    var cell = row.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellAddress);
                    if (cell == null)
                    {
                        cell = new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(map.HeaderName),
                            CellReference = cellAddress
                        };
                        if (!createdSheet)
                        {
                            if (!styleIndex.HasValue)
                            {
                                //default it, so we don't look again if we cant find a prior cell or the row does not have a style index
                                styleIndex = 0;
                                //try to get the row style index if we have one
                                if (row.StyleIndex != null && row.StyleIndex > 0)
                                {
                                    styleIndex = row.StyleIndex;
                                }
                                else
                                {
                                    //try to copy the style from a prior cell that has a style if we have one
                                    var priorCell = row.Descendants<Cell>().LastOrDefault(c => c.StyleIndex != null);
                                    if (priorCell != null && priorCell.StyleIndex != null && priorCell.StyleIndex > 0)
                                    {
                                        styleIndex = priorCell.StyleIndex;
                                    }
                                }
                            }
                            // ReSharper disable once PossibleInvalidOperationException
                            cell.StyleIndex = styleIndex.Value;
                        }
                        row.AppendChild(cell);
                    }
                    else if (map.UpdateHeader) //only update the header if we are told to
                    {
                        //overwrite the cell header value
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(map.HeaderName);
                    }
                }
            }
            if (isNewHeaderRow)
            {
                //write the row to the sheet data
                row.RowIndex = (uint)this.HeaderRow;
                sheetData.AppendChild(row);
            }
        }


        /// <summary>
        /// Creates the new sheet and adds it to the workbook part.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>WorksheetPart.</returns>
        /// <exception cref="System.FormatException">Sheet name is too long. Excel sheet names have to be 30 characters or less.</exception>
        private WorksheetPart CreateNewSheet(WorkbookPart workbookPart, string sheetName)
        {
            //sheet names can not exceed 31 characters else it WILL corrupt the sheet.
            if (sheetName.Length >= 31)
            {
                throw new FormatException("Sheet name is too long. Excel sheet names have to be 30 characters or less.");
            }

            //add the sheet if they want us to
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheetId = 1U;

            if (sheets != null)
            {
                if (sheets.Count() > 0)
                {
                    sheetId = Math.Max(sheets.Descendants<Sheet>().Select(s => s.SheetId.Value).Max() + 1U, 1U);
                }
            }
            else
            {
                sheets = new Sheets();
                // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                workbookPart.Workbook.Append(sheets);
            }
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            sheets.Append(new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = sheetName
            });

            /* TDC: This Package.Flush() line is CRAZY important. IF the package is not flushed, this new sheet
			 * will NOT LOAD when writing to a memorystream and the workbook will be corrupted. TRUST ME.
			 * Oddly, this line is NOT needed if writing to a file. */
            _spreadSheet.Package.Flush();

            return worksheetPart;
        }

        /// <summary>
        /// Gets the headers from the properties of the type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>IDictionary&lt;System.String, System.String&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">type</exception>
        protected virtual IDictionary<string, string> GetHeadersFromType(Type type)
        {
            if (type == null) { throw new ArgumentNullException("type"); }

            var properties = type.GetProperties().Where(pi => pi.GetGetMethod() != null);
            return GetHeadersFromType(properties);
        }

        /// <summary>
        /// Gets the headers from the properties.
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <returns>IDictionary&lt;System.String, System.String&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">properties</exception>
        protected virtual IDictionary<string, string> GetHeadersFromType(IEnumerable<PropertyInfo> properties)
        {
            if (properties == null) { throw new ArgumentNullException("properties"); }

            var i = 1;
            IDictionary<string, string> headers = properties.ToDictionary(key => FileIOUtilities.ConvertExcelNumberToLetter(i++), val => val.Name, comparer);
            return headers;
        }

        /// <summary>
        /// Gets the headers.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="headerRowIndex">Index of the header row.</param>
        /// <returns>IDictionary&lt;System.String, System.String&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">workbookPart
        /// or
        /// sheet
        /// or
        /// stringTable</exception>
        /// <exception cref="System.ArgumentException">headerRowIndex must be greater than or equal to 1.</exception>
        /// <exception cref="FileFormatException">
        /// The header row was not found in the file.
        /// or
        /// The header row does not contain any cells.
        /// </exception>
        /// <exception cref="System.IO.FileFormatException">workbookPart
        /// or
        /// sheet
        /// or
        /// stringTable</exception>
        protected virtual IDictionary<string, string> GetHeaders(WorkbookPart workbookPart, SheetData sheetData, int headerRowIndex)
        {
            if (workbookPart == null) { throw new ArgumentNullException("workbookPart"); }
            if (sheetData == null) { throw new ArgumentNullException("sheetData"); }
            if (headerRowIndex <= 0) { throw new ArgumentException("headerRowIndex must be greater than or equal to 1."); }

            //make a case insensitive dictionary
            IDictionary<string, string> headers = new Dictionary<string, string>(comparer);

            //get the shared string table
            var shareStringPart = GetSharedStringPart(workbookPart);
            var stringTable = shareStringPart.SharedStringTable;
            var headerRow = sheetData.Elements<Row>().ElementAtOrDefault(headerRowIndex - 1); //excel index is 1 based so subtract one

            if (headerRow == null)
            {
                headerRow = new Row() { RowIndex = Convert.ToUInt32(headerRowIndex) };
                sheetData.AppendChild(headerRow);
                return headers;
                //throw new FileFormatException("The header row was not found in the file.");
            }

            if (headerRow.HasChildren)
            {
                //loop through all of the cells in the row, building a list of header keys.
                foreach (var cell in headerRow.Descendants<Cell>())
                {
                    //get the column name minus the row number
                    var cellKey = FileIOUtilities.GetColumnKey(cell.CellReference.Value);
                    var value = string.Empty;

                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                    {
                        //read the text value out of the shared string table.
                        value = stringTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText;
                    }
                    else if (cell.CellValue != null)
                    {
                        value = cell.CellValue.Text;
                    }

                    headers.Add(cellKey, value);
                }
            }
            else
            {
                throw new FileFormatException("The header row does not contain any cells.");
            }
            //remove any characters from the headers that are not valid for .net property names
            headers = FileIOUtilities.FixupHeaders(headers);
            return headers;
        }

        /// <summary>
        /// Gets the shared string part.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <returns>SharedStringTablePart.</returns>
        /// <exception cref="System.ArgumentNullException">workbookPart</exception>
        protected virtual SharedStringTablePart GetSharedStringPart(WorkbookPart workbookPart)
        {
            if (workbookPart == null) { throw new ArgumentNullException("workbookPart"); }

            var shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (shareStringPart == null)
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                shareStringPart.SharedStringTable = new SharedStringTable();
                workbookPart.Workbook.Save();
            }
            return shareStringPart;
        }

        /// <summary>
        /// Gets the type of the excel cell type from property.
        /// </summary>
        /// <param name="propertyType">Type of the property.</param>
        /// <returns>ExcelCellType.</returns>
        private static ExcelCellType GetExcelCellTypeFromPropertyType(Type propertyType)
        {
            var notNullType = Nullable.GetUnderlyingType(propertyType);
            if (notNullType != null) { propertyType = notNullType; }

            ExcelCellType excelCellType;
            switch (Type.GetTypeCode(propertyType))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    excelCellType = ExcelCellType.Number;
                    break;
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    excelCellType = ExcelCellType.Decimal;
                    break;
                case TypeCode.DateTime:
                    excelCellType = ExcelCellType.Date;
                    break;
                // ReSharper disable once RedundantCaseLabel
                case TypeCode.Boolean:
                default:
                    excelCellType = ExcelCellType.General;
                    break;
            }
            return excelCellType;
        }

        /// <summary>
        /// Gets the cell type from property.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>CellValues.</returns>
        protected virtual CellValues GetCellTypeFromType(Type type)
        {
            var notNullType = Nullable.GetUnderlyingType(type);
            if (notNullType != null) { type = notNullType; }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return CellValues.Number;
                //case TypeCode.Boolean: // will cause sheet corruption! yay! cause we love us some sheet corruption! so, do not use.
                //	return CellValues.Boolean;
                case TypeCode.DateTime:
                    return CellValues.Date;
                default:
                    return CellValues.String;
            }
        }

        /// <summary>
        /// Parses the lambda string <paramref name="expression" /> into a <see cref="LambdaExpression" />.
        /// </summary>
        /// <typeparam name="TInput">The type of the t input.</typeparam>
        /// <typeparam name="TResult">The type of the t result.</typeparam>
        /// <param name="expression">The expression.</param>
        /// <param name="inputParameter">The input parameter.</param>
        /// <returns>Expression&lt;Func&lt;TInput, TResult&gt;&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">expression</exception>
        protected virtual Expression<Func<TInput, TResult>> ParseLambda<TInput, TResult>(string expression, ParameterExpression inputParameter)
        {
            if (string.IsNullOrEmpty(expression)) throw new ArgumentNullException("expression");

            //string paramString = expression.Substring(0, expression.IndexOf("=>")).Trim().Trim('(', ')').Trim();
            var lambdaString = expression.Substring(expression.IndexOf("=>", StringComparison.Ordinal) + 2).Trim();
            //ParameterExpression param = Expression.Parameter(typeof(TInput), paramString);
            return (Expression<Func<TInput, TResult>>)DL.DynamicExpression.ParseLambda(new[] { inputParameter }, typeof(TResult), lambdaString, null);
        }

        /// <summary>
        /// Gets the index of the format.
        /// </summary>
        /// <param name="cellType">Type of the cell.</param>
        /// <param name="formatCode">The format code.</param>
        /// <returns>System.UInt32.</returns>
        private uint GetFormatIndex(ExcelCellType cellType, string formatCode)
        {
            //try to find the custom format that matches one added prior
            var format = _customFormats.FirstOrDefault(x => x.Key.CellType == cellType && x.Key.FormatCode == formatCode);

            //no specific format found, so get the default for this cell type
            if (format.Key == null)
            {
                format = _customFormats.FirstOrDefault(x => x.Key.CellType == cellType && x.Key.IsDefault);
            }

            //could not find a style or a default
            if (format.Key == null) { return 0U; }

            return format.Value;
        }

        /// <summary>
        /// Adds the workbook styles.
        /// </summary>
        /// <param name="workBookPart">The work book part.</param>
        /// <exception cref="System.ArgumentNullException">workBookPart</exception>
        [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "workBook")]
        [SuppressMessage("ReSharper", "RedundantAssignment")]
        protected virtual void AddWorkbookStyles(WorkbookPart workBookPart)
        {
            if (workBookPart == null) { throw new ArgumentNullException("workBookPart"); }

            var maxNumberFormatId = 165;
            //add the built in styles
            var stylesPart = workBookPart.WorkbookStylesPart;
            if (stylesPart == null) { stylesPart = workBookPart.AddNewPart<WorkbookStylesPart>(); }
            if (stylesPart.Stylesheet == null) { stylesPart.Stylesheet = CreateStyleSheet(); }
            if (stylesPart.Stylesheet.CellFormats == null) { stylesPart.Stylesheet.CellFormats = new CellFormats(); }
            if (stylesPart.Stylesheet.NumberingFormats == null) { stylesPart.Stylesheet.NumberingFormats = new NumberingFormats(); }

            if (stylesPart.Stylesheet.NumberingFormats.Count() > 0)
            {
                maxNumberFormatId = GetMaxNumberFormatId(stylesPart);
            }

            var ci = CultureInfo.CurrentCulture;
            var numFormat = ci.NumberFormat;

            AddNumberFormat(stylesPart, ExcelCellType.Date, ci.DateTimeFormat.ShortDatePattern, maxNumberFormatId++, true);
            AddNumberFormat(stylesPart, ExcelCellType.Number, "0", maxNumberFormatId++, true);
            AddNumberFormat(stylesPart, ExcelCellType.Decimal, string.Format("0{0}{1}", numFormat.NumberDecimalSeparator, new string('0', numFormat.NumberDecimalDigits)), maxNumberFormatId++, true);
            AddNumberFormat(stylesPart, ExcelCellType.Percent, "0" + numFormat.PercentSymbol, maxNumberFormatId++, true);
            AddNumberFormat(stylesPart, ExcelCellType.Currency, GetExcelCurrencyFormat(numFormat), maxNumberFormatId++, true);
        }

        /// <summary>
        /// Adds the number format.
        /// </summary>
        /// <param name="stylesPart">The styles part.</param>
        /// <param name="cellType">Type of the cell.</param>
        /// <param name="formatCode">The format code.</param>
        /// <param name="numberFormatId">The number format identifier.</param>
        /// <param name="isDefault">if set to <c>true</c> [is default].</param>
        /// <exception cref="System.ArgumentNullException">stylesPart</exception>
        protected virtual void AddNumberFormat(WorkbookStylesPart stylesPart, ExcelCellType cellType, string formatCode, int numberFormatId = 0, bool isDefault = false)
        {
            if (stylesPart == null) { throw new ArgumentNullException("stylesPart"); }

            var customFormat = _customFormats.FirstOrDefault(x => x.Key.CellType == cellType && x.Key.FormatCode == formatCode);
            //we already have added a format that matches
            NumberingFormat format;

            if (customFormat.Key == null)
            {
                //if they supplied us with a numberformatid use it. it needs to be unique however
                if (stylesPart.Stylesheet.NumberingFormats.Count() > 0 && numberFormatId == 0)
                {
                    numberFormatId = GetMaxNumberFormatId(stylesPart);
                }

                format = new NumberingFormat()
                {
                    NumberFormatId = Convert.ToUInt32(numberFormatId),
                    FormatCode = StringValue.FromString(formatCode)
                };

                //add our new numberformat
                stylesPart.Stylesheet.NumberingFormats.AppendChild(format);
            }
            else
            {
                format = stylesPart.Stylesheet.NumberingFormats.Descendants<NumberingFormat>().FirstOrDefault(x => x.NumberFormatId != null && x.NumberFormatId == customFormat.Key.NumberFormatId);
            }

            // ReSharper disable once PossibleNullReferenceException
            AddCellFormat(stylesPart, cellType, isDefault, Convert.ToInt32(format.NumberFormatId.Value), format.FormatCode);
        }

        /// <summary>
        /// Adds the cell format.
        /// </summary>
        /// <param name="stylesPart">The styles part.</param>
        /// <param name="cellType">Type of the cell.</param>
        /// <param name="isDefault">if set to <c>true</c> [is default].</param>
        /// <param name="numberFormatId">The number format identifier.</param>
        /// <param name="formatCode">The format code.</param>
        /// <exception cref="System.ArgumentNullException">stylesPart</exception>
        protected virtual void AddCellFormat(WorkbookStylesPart stylesPart, ExcelCellType cellType, bool isDefault, int numberFormatId = 0, string formatCode = null)
        {
            if (stylesPart == null) { throw new ArgumentNullException("stylesPart"); }

            var numFormatId = (uint)numberFormatId;

            var cellFormat = stylesPart.Stylesheet.CellFormats.Descendants<CellFormat>().FirstOrDefault(x => (x.NumberFormatId != null && x.NumberFormatId == numFormatId));
            if (cellFormat != null) { return; }

            cellFormat = new CellFormat()
            {
                FormatId = 0,
                FillId = 0,
                BorderId = 0,
                NumberFormatId = numFormatId,
                QuotePrefix = BooleanValue.FromBoolean(cellType == ExcelCellType.Quoted),
                ApplyNumberFormat = BooleanValue.FromBoolean(numFormatId >= 1U)
            };

            //add the cellformat that goes to this cellformat
            stylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            //build our formatkey and add it to our custom formats
            var key = new FormatsKey()
            {
                CellType = cellType,
                FormatCode = formatCode,
                IsDefault = isDefault,
                NumberFormatId = numberFormatId
            };
            SaveStyles(stylesPart);
            _customFormats.Add(key, Convert.ToUInt32(stylesPart.Stylesheet.CellFormats.Count() - 1));
        }

        /// <summary>
        /// Saves the styles.
        /// </summary>
        /// <param name="stylesPart">The styles part.</param>
        /// <exception cref="System.ArgumentNullException">stylesPart</exception>
        protected virtual void SaveStyles(WorkbookStylesPart stylesPart)
        {
            if (stylesPart == null) { throw new ArgumentNullException("stylesPart"); }

            //update the counts as they don't seem to automatically update otherwise. 
            stylesPart.Stylesheet.NumberingFormats.Count = (uint)stylesPart.Stylesheet.NumberingFormats.ChildElements.Count();
            stylesPart.Stylesheet.CellFormats.Count = (uint)stylesPart.Stylesheet.CellFormats.ChildElements.Count();
            stylesPart.Stylesheet.Save();
            _spreadSheet.Package.Flush(); //just cya :|
        }

        /// <summary>
        /// Gets the excel currency format.
        /// </summary>
        /// <param name="numFormat">The number format.</param>
        /// <returns>System.String.</returns>
        private static string GetExcelCurrencyFormat(NumberFormatInfo numFormat)
        {
            //https://msdn.microsoft.com/en-us/library/system.globalization.numberformatinfo.currencynegativepattern(v=vs.110).aspx
            string[] currencyNegativePatterns = { "(\"{0}\"{1})", "-\"{0}\"{1}", "\"{0}\"-{1}", "\"{0}\"{1}-", "({1}\"{0}\")",
                                 "-{1}\"{0}\"", "{1}-\"{0}\"", "{1}\"{0}\"-", "-{1} \"{0}\"", "-\"{0}\" {1}",
                                 "{1} \"{0}\"-", "\"{0}\" {1}-", "\"{0}\" -{1}", "{1}- \"{0}\"", "(\"{0}\" {1})",
                                 "({1} \"{0}\")" };

            //setup the currency format per the region settings...
            var decimals = new string('0', numFormat.CurrencyDecimalDigits);
            var baseCurrencyFormat = string.Format("#{0}##0{1}{2}", numFormat.CurrencyGroupSeparator, numFormat.CurrencyDecimalSeparator, decimals);
            var positiveCurrencyFormat = string.Format("\"{0}\"{1};", numFormat.CurrencySymbol, baseCurrencyFormat);
            var negativeCurrencyFormat = string.Format(currencyNegativePatterns[numFormat.CurrencyNegativePattern] + ";", numFormat.CurrencySymbol, baseCurrencyFormat);
            return positiveCurrencyFormat + negativeCurrencyFormat;

        }

        /// <summary>
        /// Gets the maximum number format identifier.
        /// </summary>
        /// <param name="stylesPart">The styles part.</param>
        /// <returns>System.Int32.</returns>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0")]
        protected virtual int GetMaxNumberFormatId(WorkbookStylesPart stylesPart)
        {
            //use 165 as the base number id, as the built in ones go from 1 to 164
            var numberFormatId = Math.Max(Convert.ToInt32(stylesPart.Stylesheet.NumberingFormats.Descendants<NumberingFormat>().Select(nf => nf.NumberFormatId.Value).Max() + 1), 165);
            return numberFormatId;
        }

        /// <summary>
        /// Creates the style sheet.
        /// </summary>
        /// <returns>Stylesheet.</returns>
        [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
        protected virtual Stylesheet CreateStyleSheet()
        {
            //everything here is needed at a minimum for the styles to not corrupt the sheet with a new workbook
            var styleSheet = new Stylesheet();
            styleSheet.Fonts = new Fonts() { Count = 1, KnownFonts = BooleanValue.FromBoolean(true) };
            styleSheet.Fonts.Append(new Font()
            {
                FontSize = new FontSize() { Val = 11 },
                FontName = new FontName() { Val = "Calibri" },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(FontSchemeValues.Minor) }
            });
            styleSheet.Fills = new Fills() { Count = 1 };
            styleSheet.Fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.None) }
            });
            styleSheet.Borders = new Borders() { Count = 1 };
            styleSheet.Borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            styleSheet.CellFormats = new CellFormats() { Count = 1 };
            styleSheet.CellFormats.Append(new CellFormat());
            styleSheet.NumberingFormats = new NumberingFormats() { Count = 0 };

            return styleSheet;
        }

        #endregion	helpers

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA1816:CallGCSuppressFinalizeCorrectly")]
        public void Dispose()
        {
            Close();
            if (this.SpreadsheetStream != null)
            {
                this.SpreadsheetStream.Dispose();
                this.SpreadsheetStream = null;
            }
        }

        /// <summary>
        /// Closes this writer instance and the underlying spread sheet.
        /// </summary>
        public virtual void Close()
        {
            _isClosed = true;
            if (_spreadSheet != null)
            {
                //just doing a save for cya sake
                _spreadSheet.Package.Flush();
                _spreadSheet.WorkbookPart.Workbook.Save();
                _spreadSheet.Dispose();
                _spreadSheet = null;
            }
            if (_spreadsheetPackage != null)
            {
                _spreadsheetPackage.Close();
                _spreadsheetPackage = null;
            }
        }
        #endregion
    }
}
