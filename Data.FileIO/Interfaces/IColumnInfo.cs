// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IColumnInfo.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Linq.Expressions;
namespace Data.FileIO.Interfaces
{
	/// <summary>
	/// Interface IColumnInfo
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="Data.FileIO.Interfaces.IColumnInfoBase" />
	public interface IColumnInfo<T> : IColumnInfoBase
	{
		/// <summary>
		/// Gets the value function. The value function will be used to provide the value for the cell. The row object being processed will be the parameter type.
		/// </summary>
		/// <value>The value function.</value>
		Func<T, object> ValueFunction { get; }

		/// <summary>
		/// Gets the type of the value function.
		/// </summary>
		/// <value>The type of the value function.</value>
		Type ValueFunctionType { get; }
	}
}
