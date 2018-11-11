// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IObjectValidator.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
namespace Data.FileIO.Common.Interfaces
{
	/// <summary>
	/// Interface IObjectValidator
	/// </summary>
	public interface IObjectValidator
	{
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
		T? GetRowValue<T>(dynamic row, string propertyName, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false) where T : struct;
		/// <summary>
		/// Gets the row value.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="valueFunction">The value function.</param>
		/// <param name="parseErrors">The parse errors.</param>
		/// <param name="parseErrorMessage">The parse error message.</param>
		/// <param name="isNullable">if set to <c>true</c> [is nullable].</param>
		/// <returns>System.Nullable&lt;T&gt;.</returns>
		T? GetRowValue<T>(Expression<Func<object>> valueFunction, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false) where T : struct;
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
		T? GetRowValue<T>(string input, ref List<string> parseErrors, string parseErrorMessage = null, bool isNullable = false, string propertyName = null) where T : struct;
		/// <summary>
		/// Tries the convert.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="input">The input.</param>
		/// <param name="newValue">The new value.</param>
		/// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
		bool TryConvert<T>(object input, out T? newValue) where T : struct;
		/// <summary>
		/// Tries the validate.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="validationObject">The validation object.</param>
		/// <param name="errors">The errors.</param>
		/// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
		bool TryValidate<T>(T validationObject, ref List<string> errors);
	}
}
