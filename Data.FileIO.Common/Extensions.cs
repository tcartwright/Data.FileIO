// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ListExtensions.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Data.FileIO.Common
{
	/// <summary>
	/// Class ListExtensions.
	/// </summary>
	public static class Extensions
	{
		/// <summary>
		/// Determines whether this instance is empty.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="list">The list.</param>
		/// <returns><c>true</c> if the specified list is empty; otherwise, <c>false</c>.</returns>
		public static bool IsEmpty<T>(this IList<T> list)
		{
			return list == null || list.Count == 0;
		}

		/// <summary>
		/// Adds or updates the value in the dictionary.
		/// </summary>
		/// <typeparam name="TKey">The type of the key.</typeparam>
		/// <typeparam name="TValue">The type of the value.</typeparam>
		/// <param name="dictionary">The dictionary.</param>
		/// <param name="key">The key.</param>
		/// <param name="value">The value.</param>
		/// <exception cref="System.ArgumentNullException">dictionary</exception>
		public static void AddOrUpdate<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue value)
		{
			if (dictionary == null) throw new ArgumentNullException("dictionary");

			if (dictionary.ContainsKey(key))
			{
				dictionary[key] = value;
			}
			else
			{
				dictionary.Add(key, value);
			}
		}

		/// <summary>
		/// Determines whether this instance is empty.
		/// </summary>
		/// <typeparam name="TKey">The type of the key.</typeparam>
		/// <typeparam name="TValue">The type of the value.</typeparam>
		/// <param name="dictionary">The dictionary.</param>
		/// <returns><c>true</c> if the specified dictionary is empty; otherwise, <c>false</c>.</returns>
		public static bool IsEmpty<TKey, TValue>(this IDictionary<TKey, TValue> dictionary)
		{
			return dictionary == null || dictionary.Count == 0;
		}
	}
}
