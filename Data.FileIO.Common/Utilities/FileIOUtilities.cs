// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ParserUtilities.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;
using Data.FileIO.Common.Interfaces;


namespace Data.FileIO.Common.Utilities
{
    /// <summary>
    /// Class ParserUtilities.
    /// </summary>
    public sealed class FileIOUtilities
    {
        /// <summary>
        /// Prevents a default instance of the <see cref="FileIOUtilities"/> class from being created.
        /// </summary>
        private FileIOUtilities()
        {

        }
        /// <summary>
        /// returns the column key of the excel cell address minus numbers.
        /// </summary>
        /// <param name="address">The address.</param>
        /// <returns>System.String.</returns>
        public static string GetColumnKey(string address)
        {
            return Regex.Replace(address, "[^A-Za-z]", "");
        }

        /// <summary>
        /// Converts the excel column name to number.
        /// </summary>
        /// <param name="columnName">Name of the column.</param>
        /// <returns>System.Int32.</returns>
        /// <exception cref="System.ArgumentNullException">columnName</exception>
        public static int ConvertExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return -1;

            columnName = GetColumnKey(columnName).ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        /// <summary>
        /// Converts the excel number to letter.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        /// <returns>System.String.</returns>
        public static string ConvertExcelNumberToLetter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        /// <summary>
        /// Fixes up the header.
        /// </summary>
        /// <param name="header">The header.</param>
        /// <returns>System.String.</returns>
        public static string FixupHeader(string header)
        {
            return Regex.Replace(header, "[^A-Za-z0-9_]*", String.Empty);
        }

        /// <summary>
        /// Turn csv headers into proper .net property names.
        /// </summary>
        /// <param name="headers">The headers.</param>
        /// <returns>System.String[].</returns>
        /// <exception cref="System.ArgumentNullException">headers</exception>
        public static string[] FixupHeaders(string[] headers)
        {
            if (headers == null) { throw new ArgumentNullException("headers"); }
            for (int i = 0; i < headers.Length; i++)
            {
                headers[i] = FixupHeader(headers[i]);
                if (String.IsNullOrWhiteSpace(headers[i]))
                {
                    headers[i] = "Col" + i.ToString();
                }
            }
            return headers;
        }

        /// <summary>
        /// Turn excel headers into proper .net property names.
        /// </summary>
        /// <param name="headers">The headers.</param>
        /// <returns>Dictionary&lt;System.String, System.String&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">headers</exception>
        public static IDictionary<string, string> FixupHeaders(IDictionary<string, string> headers)
        {
            if (headers == null) { throw new ArgumentNullException("headers"); }
            for (int i = 0; i < headers.Keys.Count(); i++)
            {
                string key = headers.Keys.ElementAt(i);
                if (!String.IsNullOrWhiteSpace(headers[key]))
                {
                    headers[key] = FixupHeader(headers[key]);
                }
                else
                {
                    //as the header row may contain headers with no name, assume cell address for property name at that point
                    headers[key] = key;
                }
            }
            return headers;
        }

        /// <summary>
        /// Converts the headers and values into a dynamic object.
        /// </summary>
        /// <param name="headers">The headers.</param>
        /// <param name="values">The values.</param>
        /// <param name="rowIndex">The rowIndex.</param>
        /// <returns>dynamic.</returns>
        /// <exception cref="System.ArgumentNullException">
        /// headers
        /// or
        /// values
        /// </exception>
        public static dynamic RowToExpando(IDictionary<string, string> headers, IDictionary<string, string> values, int rowIndex = -1)
        {
            if (headers == null) { throw new ArgumentNullException("headers"); }
            if (values == null) { throw new ArgumentNullException("values"); }
            //As not every excel cell may have a value, we have to emit only those cells whose key match up to the headers
            CustomExpandoObject expandoObject = new CustomExpandoObject(true);
            foreach (string key in headers.Keys)
            {
                if (values.ContainsKey(key))
                {
                    if (!expandoObject.ContainsKey(headers[key]))
                    {
                        expandoObject.Add(headers[key], values[key]);
                    }
                    else
                    {
                        expandoObject.Add(key + "_" + headers[key], values[key]);
                    }
                }
                else
                {
                    if (expandoObject.ContainsKey(headers[key]))
                    {
                        throw new DuplicateColumnNameException(string.Format("Column '{0}' found multiple times in the excel file.", headers[key]));
                    }
                    expandoObject.Add(headers[key], String.Empty);
                }
            }

            if (rowIndex >= 0)
            {
                expandoObject.Add("__RowIndex", rowIndex);
            }

            return expandoObject;
        }


        /// <summary>
        /// Converts the headers and values into a dynamic object.
        /// </summary>
        /// <param name="headers">The headers.</param>
        /// <param name="values">The values.</param>
        /// <returns>dynamic.</returns>
        /// <exception cref="System.ArgumentNullException">
        /// headers
        /// or
        /// values
        /// </exception>
        public static dynamic RowToExpando(string[] headers, string[] values)
        {
            if (headers == null) { throw new ArgumentNullException("headers"); }
            if (values == null) { throw new ArgumentNullException("values"); }
            CustomExpandoObject expandoObject = new CustomExpandoObject(true);
            for (var i = 0; i < headers.Length; i++)
            {
                expandoObject.Add(headers[i], values[i]);
            }
            return expandoObject;
        }

        /// <summary>
        /// Maps the object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row">The row.</param>
        /// <param name="validator">The validator.</param>
        /// <param name="errors">The errors.</param>
        /// <returns>T.</returns>
        /// <exception cref="System.ArgumentNullException">
        /// row
        /// or
        /// validator
        /// or
        /// errors
        /// </exception>
        public static T MapObject<T>(dynamic row, IObjectValidator validator, ref List<string> errors)
        {
            if (row == null) { throw new ArgumentNullException("row"); }
            if (validator == null) { throw new ArgumentNullException("validator"); }
            if (errors == null) { throw new ArgumentNullException("errors"); }
            return MapObject<T>(row, -1, validator, null, ref errors);
        }

        /// <summary>
        /// Maps the object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row">The row.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="validator">The validator.</param>
        /// <param name="mapper">The mapper.</param>
        /// <param name="errors">The errors.</param>
        /// <returns>T.</returns>
        /// <exception cref="System.ArgumentNullException">
        /// row
        /// or
        /// validator
        /// or
        /// errors
        /// </exception>
        public static T MapObject<T>(dynamic row, int rowIndex, IObjectValidator validator, FileValuesMap<T> mapper, ref List<string> errors)
        {
            if (row == null) { throw new ArgumentNullException("row"); }
            if (validator == null) { throw new ArgumentNullException("validator"); }
            if (errors == null) { throw new ArgumentNullException("errors"); }

            //if the object is not a dynamic object we do not need to recreate it, just use as is
            if (!FileIOHelpers.IsDynamicType(row)) { return row; }
            T rowObj = (T)Activator.CreateInstance(typeof(T), true);

            IFileRowMapper mapperObj = null;
            //a custom mapper. will take priority if exists 
            if (mapper != null)
            {
                mapper.Invoke(ref rowObj, rowIndex, row, validator, ref errors);
            }
            else if ((mapperObj = rowObj as IFileRowMapper) != null)
            {
                mapperObj.MapValues(rowIndex, row, validator, ref errors);
            }
            else
            {
                //last ditch effort to map the values to the object if neither of the main mappers gets called
                TryMapValues<T>(rowObj, row, validator, ref errors);
            }

            return rowObj;
        }

        /// <summary>
        /// Tries to map values from the row to the object using reflection.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rowObj">The new object.</param>
        /// <param name="row">The row.</param>
        /// <param name="validator">The validator.</param>
        /// <param name="errors">The errors.</param>
        /// <exception cref="System.ArgumentNullException">
        /// rowObj
        /// or
        /// row
        /// or
        /// validator
        /// or
        /// errors
        /// </exception>
        public static void TryMapValues<T>(T rowObj, dynamic row, IObjectValidator validator, ref List<string> errors)
        {
            if (rowObj == null) { throw new ArgumentNullException("rowObj"); }
            if (row == null) { throw new ArgumentNullException("row"); }
            if (validator == null) { throw new ArgumentNullException("validator"); }
            if (errors == null) { throw new ArgumentNullException("errors"); }

            CustomExpandoObject expando = row as CustomExpandoObject;
            if (expando == null)
            {
                expando = new CustomExpandoObject(row, true);
            }

            var getRowValueMethod = validator.GetType().GetMethods().First(x => x.Name == "GetRowValue" && x.GetParameters().Count() >= 5 && x.GetParameters().First().ParameterType == typeof(string));
            PropertyInfo[] properties = typeof(T).GetProperties();

            foreach (var prop in properties)
            {
                if (expando.ContainsKey(prop.Name)) //there may be properties that were not imported
                {
                    string val = Convert.ToString(expando[prop.Name]);

                    Type nonNullableType = FileIOHelpers.GetNonNullableType(prop.PropertyType);
                    if (nonNullableType.IsValueType)
                    {
                        bool isNullable = Nullable.GetUnderlyingType(prop.PropertyType) != null;
                        //convert the Helpers.GetValue method into a generic method of the same type as the property
                        var genericGetvalueMethod = getRowValueMethod.MakeGenericMethod(new Type[] { nonNullableType });
                        var value = genericGetvalueMethod.Invoke(validator, new object[] { val, errors, null, isNullable, prop.Name });
                        prop.SetValue(rowObj, value, null); //this should use the setrowvalue that outputs conversion errors
                    }
                    else
                    {
                        prop.SetValue(rowObj, Convert.ChangeType(val, prop.PropertyType), null);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the return type of the expression.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns>Type.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
        public static Type GetExpressionType<T>(Expression<Func<T, object>> expression)
        {
            if (expression == null) { throw new ArgumentNullException("expression"); }
            if (expression.Body.NodeType == ExpressionType.MemberAccess)
            {
                var expr = expression.Body as MemberExpression;
                if (expr != null) { return expr.Type; }
            }
            else if (expression.Body.NodeType == ExpressionType.Convert || expression.Body.NodeType == ExpressionType.ConvertChecked)
            {
                var expr = expression.Body as UnaryExpression;
                if (expr != null) { return expr.Operand.Type; }
            }
            else if (expression.Body.NodeType == ExpressionType.Call)
            {
                var expr = expression.Body as MethodCallExpression;
                if (expr != null) { return expr.Method.ReturnType; }
            }
            return expression.Body.Type;
        }
    }
}
