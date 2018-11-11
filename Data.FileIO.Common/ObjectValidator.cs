// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ObjectValidator.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;

namespace Data.FileIO.Common
{
	/// <summary>
	/// Class ObjectValidator.
	/// </summary>
	public class ObjectValidator : IObjectValidator
	{
		/// <summary>
		/// Will try to validate the objects and return a list of strings that match the error condition.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="validationObject">The validation object.</param>
		/// <param name="errors">The errors.</param>
		/// <returns><c>true</c> if Validation passed, <c>false</c> otherwise.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// validationObject
		/// or
		/// errors
		/// </exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1")]
		public virtual bool TryValidate<T>(T validationObject, ref List<string> errors)
		{
			if (validationObject == null) { throw new ArgumentNullException("validationObject"); }
			if (errors == null) { throw new ArgumentNullException("errors"); }

			var context = new ValidationContext(validationObject, serviceProvider: null, items: null);
			ICollection<ValidationResult> validationResults = new List<ValidationResult>();
			Validator.TryValidateObject(validationObject, context, validationResults, validateAllProperties: true);

			var results = new List<string>(validationResults.Select(x => x.ErrorMessage));

			if (validationObject is IValidatableObject)
			{
				var iValidatableObjectErrors = ((IValidatableObject)validationObject).Validate(context);
				results.AddRange(iValidatableObjectErrors.Select(x => x.ErrorMessage));
			}

			//get rid of duplicate errors
			results = new List<string>(results.Distinct());

			if (results != null)
			{
				errors.AddRange(results);
			}

			return results.Count == 0;
		}

		/// <summary>
		/// Gets the row value.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="valueFunction">The row value.</param>
		/// <param name="parseErrors">The parse errors.</param>
		/// <param name="parseErrorMessage">The parse error message.</param>
		/// <param name="isNullable">if set to <c>true</c> [is nullable].</param>
		/// <returns>System.Nullable&lt;T&gt;.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// valueFunction
		/// or
		/// parseErrors
		/// </exception>
		public virtual T? GetRowValue<T>(Expression<Func<object>> valueFunction, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false) where T : struct
		{
			if (valueFunction == null) { throw new ArgumentNullException("valueFunction"); }
			if (parseErrors == null) { throw new ArgumentNullException("parseErrors"); }

			string input = Convert.ToString(valueFunction.Compile().Invoke());

			var propertyExpression = valueFunction.Body as MemberExpression;
			string propertyName = propertyExpression.Member.Name;

			return GetRowValue<T>(input, ref parseErrors, parseErrorMessage, isNullable, propertyName);
		}

		/// <summary>
		/// Gets the row value.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="row">The row.</param>
		/// <param name="propertyName">Name of the property.</param>
		/// <param name="parseErrors">The parse errors.</param>
		/// <param name="parseErrorMessage">The parse error message.</param>
		/// <param name="isNullable">if set to <c>true</c> [is nullable].</param>
		/// <returns>System.Nullable&lt;T&gt;.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// row
		/// or
		/// propertyName
		/// or
		/// parseErrors
		/// </exception>
		public virtual T? GetRowValue<T>(dynamic row, string propertyName, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false) where T : struct
		{
			if (row == null) { throw new ArgumentNullException("row"); }
			if (String.IsNullOrWhiteSpace(propertyName)) { throw new ArgumentNullException("propertyName"); }
			if (parseErrors == null) { throw new ArgumentNullException("parseErrors"); }

			string input = null;
			bool valueFound = false;
			if (FileIOHelpers.IsDynamicType(row))
			{
				IDictionary<string, object> dict = row as IDictionary<string, object>;
				if (dict.ContainsKey(propertyName))
				{
					input = Convert.ToString(row[propertyName]);
					valueFound = true;
				}
			}
			else
			{
				//the object is not a dynamic object, so get the value using reflection
				var prop = row.GetType().GetProperty(propertyName);
				if (prop != null)
				{
					input = Convert.ToString(prop.GetValue(row, null));
					valueFound = true;
				}
			}

			if (valueFound)
			{
				return GetRowValue<T>(input, ref parseErrors, parseErrorMessage, isNullable, propertyName);
			}
			else
			{
				return default(T?);
			}
		}

		/// <summary>
		/// Gets the row value.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="input">The input.</param>
		/// <param name="parseErrors">The parse errors.</param>
		/// <param name="parseErrorMessage">The parse error message.</param>
		/// <param name="isNullable">if set to <c>true</c> [is nullable].</param>
		/// <param name="propertyName">Name of the property.</param>
		/// <returns>System.Nullable&lt;T&gt;.</returns>
		/// <exception cref="System.ArgumentNullException">parseErrors</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "1")]
		public virtual T? GetRowValue<T>(string input, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false, string propertyName = null) where T : struct
		{
			if (parseErrors == null) { throw new ArgumentNullException("parseErrors"); }

			Type type = typeof(T);
			T? retVal = default(T?);

			//if the type if a value type, and the string is empty then just return the default which should be null
			if (type.IsValueType && String.IsNullOrWhiteSpace(input))
			{
				return isNullable ? default(T?) : default(T);
			}

			//try to convert the string value to its native data type
			if (!TryConvert<T>(input, out retVal))
			{
				//the caller did not give us an error message so let us generate one.
				if (String.IsNullOrEmpty(parseErrorMessage))
				{
					parseErrorMessage = String.Format("The field {0} is not a valid {1}.", propertyName, typeof(T).Name);
				}
				//append the new error to the errors
				parseErrors.Add(parseErrorMessage);

				//return the appropriate value if the field is nullable or not.
				return isNullable ? default(T?) : default(T);
			}

			if (isNullable && typeof(T) == typeof(DateTime) && retVal.Equals(DateTime.MinValue))
			{
				retVal = null;
			}

			return retVal;
		}

		/// <summary>
		/// Tries to convert the value to the type specified.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="input">The input.</param>
		/// <param name="newValue">The newval.</param>
		/// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
		/// <exception cref="System.ArgumentNullException">input</exception>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
		public virtual bool TryConvert<T>(object input, out T? newValue) where T : struct
		{
			if (input == null) { throw new ArgumentNullException("input"); }

			Type type = typeof(T);

			var converter = TypeDescriptor.GetConverter(type);
			newValue = default(T?);
			try
			{
				//try to use the converter first as it is more efficient
				if (converter != null && converter.CanConvertFrom(input.GetType()))
				{
					newValue = (T)converter.ConvertFrom(input);
				}
				else
				{
					newValue = (T)Convert.ChangeType(input, type);
				}
				return true;
			}
			catch
			{
				return false;
			}
		}
	}
}
