// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="CustomExpando.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;

namespace Data.FileIO.Common
{
	/// <summary>
	/// Provides a case-insensitive expando, that also offers a string and int indexer
	/// </summary>
	/// <seealso cref="System.Dynamic.DynamicObject" />
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
	public class CustomExpandoObject : DynamicObject, IDictionary<string, object>
	{
		/// <summary>
		/// The _dictionary
		/// </summary>
		private Dictionary<string, object> _dictionary;
		/// <summary>
		/// The _allow missing columns
		/// </summary>
		private readonly bool _allowMissingColumns;

		/// <summary>
		/// Initializes a new instance of the <see cref="CustomExpandoObject"/> class.
		/// </summary>
		/// <param name="allowMissingColumns">if set to <c>true</c> [allow missing columns].</param>
		public CustomExpandoObject(bool allowMissingColumns = false)
		{
			_allowMissingColumns = allowMissingColumns;
			_dictionary = new Dictionary<string, object>(StringComparer.InvariantCultureIgnoreCase);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CustomExpandoObject"/> class.
		/// </summary>
		/// <param name="expando">The expando.</param>
		/// <param name="allowMissingColumns">if set to <c>true</c> [allow missing columns].</param>
		public CustomExpandoObject(IDictionary<string, object> expando, bool allowMissingColumns = false)
			: this(allowMissingColumns)
		{
			if (expando != null)
			{
				foreach (var item in expando)
				{
					_dictionary.Add(item.Key, item.Value);
				}
			}
		}

		/// <summary>
		/// Provides the implementation for operations that set member values. Classes derived from the <see cref="T:System.Dynamic.DynamicObject" /> class can override this method to specify dynamic behavior for operations such as setting a value for a property.
		/// </summary>
		/// <param name="binder">Provides information about the object that called the dynamic operation. The binder.Name property provides the name of the member to which the value is being assigned. For example, for the statement sampleObject.SampleProperty = "Test", where sampleObject is an instance of the class derived from the <see cref="T:System.Dynamic.DynamicObject" /> class, binder.Name returns "SampleProperty". The binder.IgnoreCase property specifies whether the member name is case-sensitive.</param>
		/// <param name="value">The value to set to the member. For example, for sampleObject.SampleProperty = "Test", where sampleObject is an instance of the class derived from the <see cref="T:System.Dynamic.DynamicObject" /> class, the <paramref name="value" /> is "Test".</param>
		/// <returns>true if the operation is successful; otherwise, false. If this method returns false, the run-time binder of the language determines the behavior. (In most cases, a language-specific run-time exception is thrown.)</returns>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0")]
		public override bool TrySetMember(SetMemberBinder binder, object value)
		{
			_dictionary[binder.Name] = value;
			return true;
		}

		/// <summary>
		/// Provides the implementation for operations that get member values. Classes derived from the <see cref="T:System.Dynamic.DynamicObject" /> class can override this method to specify dynamic behavior for operations such as getting a value for a property.
		/// </summary>
		/// <param name="binder">Provides information about the object that called the dynamic operation. The binder.Name property provides the name of the member on which the dynamic operation is performed. For example, for the Console.WriteLine(sampleObject.SampleProperty) statement, where sampleObject is an instance of the class derived from the <see cref="T:System.Dynamic.DynamicObject" /> class, binder.Name returns "SampleProperty". The binder.IgnoreCase property specifies whether the member name is case-sensitive.</param>
		/// <param name="result">The result of the get operation. For example, if the method is called for a property, you can assign the property value to <paramref name="result" />.</param>
		/// <returns>true if the operation is successful; otherwise, false. If this method returns false, the run-time binder of the language determines the behavior. (In most cases, a run-time exception is thrown.)</returns>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0")]
		public override bool TryGetMember(GetMemberBinder binder, out object result)
		{
			bool ret = _dictionary.TryGetValue(binder.Name, out result);
			if (_allowMissingColumns) { ret = true; }
			return ret;
		}

		/// <summary>
		/// Adds an element with the provided key and value to the <see cref="T:System.Collections.Generic.IDictionary`2" />.
		/// </summary>
		/// <param name="key">The object to use as the key of the element to add.</param>
		/// <param name="value">The object to use as the value of the element to add.</param>
		public void Add(string key, object value)
		{
			_dictionary.Add(key, value);
		}

		/// <summary>
		/// Determines whether the <see cref="T:System.Collections.Generic.IDictionary`2" /> contains an element with the specified key.
		/// </summary>
		/// <param name="key">The key to locate in the <see cref="T:System.Collections.Generic.IDictionary`2" />.</param>
		/// <returns>true if the <see cref="T:System.Collections.Generic.IDictionary`2" /> contains an element with the key; otherwise, false.</returns>
		public bool ContainsKey(string key)
		{
			return _dictionary.ContainsKey(key);
		}

		/// <summary>
		/// Gets an <see cref="T:System.Collections.Generic.ICollection`1" /> containing the keys of the <see cref="T:System.Collections.Generic.IDictionary`2" />.
		/// </summary>
		/// <value>The keys.</value>
		public ICollection<string> Keys
		{
			get { return _dictionary.Keys; }
		}

		/// <summary>
		/// Removes the element with the specified key from the <see cref="T:System.Collections.Generic.IDictionary`2" />.
		/// </summary>
		/// <param name="key">The key of the element to remove.</param>
		/// <returns>true if the element is successfully removed; otherwise, false.  This method also returns false if <paramref name="key" /> was not found in the original <see cref="T:System.Collections.Generic.IDictionary`2" />.</returns>
		public bool Remove(string key)
		{
			return _dictionary.Remove(key);
		}

		/// <summary>
		/// Gets the value associated with the specified key.
		/// </summary>
		/// <param name="key">The key whose value to get.</param>
		/// <param name="value">When this method returns, the value associated with the specified key, if the key is found; otherwise, the default value for the type of the <paramref name="value" /> parameter. This parameter is passed uninitialized.</param>
		/// <returns>true if the object that implements <see cref="T:System.Collections.Generic.IDictionary`2" /> contains an element with the specified key; otherwise, false.</returns>
		public bool TryGetValue(string key, out object value)
		{
			return _dictionary.TryGetValue(key, out value);
		}

		/// <summary>
		/// Gets an <see cref="T:System.Collections.Generic.ICollection`1" /> containing the values in the <see cref="T:System.Collections.Generic.IDictionary`2" />.
		/// </summary>
		/// <value>The values.</value>
		public ICollection<object> Values
		{
			get { return _dictionary.Values; }
		}

		/// <summary>
		/// Gets or sets the element with the specified key.
		/// </summary>
		/// <param name="key">The key.</param>
		/// <returns>System.Object.</returns>
		public object this[string key]
		{
			get { return _dictionary[key]; }
			set { _dictionary[key] = value; }
		}

		/// <summary>
		/// Gets or sets the <see cref="System.Object"/> at the specified index.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <returns>System.Object.</returns>
		public object this[int index]
		{
			get { return _dictionary.ElementAt(index); }
			set { _dictionary[_dictionary.ElementAt(index).Key] = value; }
		}

		/// <summary>
		/// Adds an item to the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <param name="item">The object to add to the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
		public void Add(KeyValuePair<string, object> item)
		{
			_dictionary.Add(item.Key, item.Value);
		}

		/// <summary>
		/// Removes all items from the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		public void Clear()
		{
			_dictionary.Clear();
		}

		/// <summary>
		/// Determines whether the <see cref="T:System.Collections.Generic.ICollection`1" /> contains a specific value.
		/// </summary>
		/// <param name="item">The object to locate in the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
		/// <returns>true if <paramref name="item" /> is found in the <see cref="T:System.Collections.Generic.ICollection`1" />; otherwise, false.</returns>
		public bool Contains(KeyValuePair<string, object> item)
		{
			return _dictionary.Contains(item);
		}

		/// <summary>
		/// Copies to.
		/// </summary>
		/// <param name="array">The array.</param>
		/// <param name="arrayIndex">Index of the array.</param>
		/// <exception cref="System.ArgumentNullException">array</exception>
		/// <exception cref="System.ArgumentOutOfRangeException">arrayIndex</exception>
		/// <exception cref="System.ArgumentException">Destination array is not large enough. Check array.Length and arrayIndex.</exception>
		public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
		{
			if (array == null) { throw new ArgumentNullException("array"); }
			if (arrayIndex < 0 || arrayIndex > array.Length) { throw new ArgumentOutOfRangeException("arrayIndex"); }
			if ((array.Length - arrayIndex) < _dictionary.Count) { throw new ArgumentException("Destination array is not large enough. Check array.Length and arrayIndex."); }

			foreach (var item in _dictionary)
			{
				array[arrayIndex++] = item;
			}
		}

		/// <summary>
		/// Gets the number of elements contained in the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <value>The count.</value>
		public int Count
		{
			get { return _dictionary.Count; }
		}

		/// <summary>
		/// Gets a value indicating whether the <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.
		/// </summary>
		/// <value><c>true</c> if this instance is read only; otherwise, <c>false</c>.</value>
		public bool IsReadOnly
		{
			get { return false; }
		}

		/// <summary>
		/// Removes the first occurrence of a specific object from the <see cref="T:System.Collections.Generic.ICollection`1" />.
		/// </summary>
		/// <param name="item">The object to remove from the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
		/// <returns>true if <paramref name="item" /> was successfully removed from the <see cref="T:System.Collections.Generic.ICollection`1" />; otherwise, false. This method also returns false if <paramref name="item" /> is not found in the original <see cref="T:System.Collections.Generic.ICollection`1" />.</returns>
		public bool Remove(KeyValuePair<string, object> item)
		{
			return _dictionary.Remove(item.Key);
		}

		/// <summary>
		/// Returns an enumerator that iterates through the collection.
		/// </summary>
		/// <returns>A <see cref="T:System.Collections.Generic.IEnumerator`1" /> that can be used to iterate through the collection.</returns>
		public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
		{
			return _dictionary.GetEnumerator();
		}

		/// <summary>
		/// Returns an enumerator that iterates through a collection.
		/// </summary>
		/// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return _dictionary.GetEnumerator();
		}
	}

}
