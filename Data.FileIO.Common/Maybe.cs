// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="Maybe.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Data.FileIO.Common
{
	/// <summary>
	/// Maybe extension method class
	/// </summary>
	public static class MaybeExtensions
	{
		/// <summary>
		/// Will return the instance if the instance is not empty, else it will return a empty enumeration.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="instance">The instance.</param>
		/// <returns>Maybe&lt;T&gt;.</returns>
		public static Maybe<T> Maybe<T>(this T instance)
		{
			return new Maybe<T>(instance);
		}
	}

	/// <summary>
	/// A maybe monad class. Will return the instance if the instance is not empty, else it will return a empty enumeration.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <seealso cref="System.Collections.Generic.IEnumerable{T}" />
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class Maybe<T> : IEnumerable<T>
	{
		/// <summary>
		/// The instance
		/// </summary>
		private readonly T instance;

		/// <summary>
		/// Initializes a new instance of the <see cref="Maybe{T}" /> class.
		/// </summary>
		/// <param name="instance">The instance.</param>
		public Maybe(T instance)
		{
			this.instance = instance;
		}

		/// <summary>
		/// Gets the value.
		/// </summary>
		/// <value>The value.</value>
		public T Value
		{
			get { return instance; }
		}

		/// <summary>
		/// Returns an enumerator that iterates through the collection.
		/// </summary>
		/// <returns>A <see cref="T:System.Collections.Generic.IEnumerator`1" /> that can be used to iterate through the collection.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			bool isEmpty = instance == null
				|| String.IsNullOrWhiteSpace(Convert.ToString(instance))
				|| (instance is IList && ((IList)instance).Count == 0); //any sort of collection, even arrays

			return Enumerable.Repeat(instance, isEmpty ? 0 : 1).GetEnumerator();
		}

		/// <summary>
		/// Returns an enumerator that iterates through a collection.
		/// </summary>
		/// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.GetEnumerator();
		}
	}

}
