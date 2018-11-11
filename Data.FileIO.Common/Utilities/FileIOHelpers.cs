// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="Helpers.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using System.Xml.Serialization;
using Data.FileIO.Common.Interfaces;

namespace Data.FileIO.Common.Utilities
{
	/// <summary>
	/// Class Helpers.
	/// </summary>
	public sealed class FileIOHelpers
	{
		/// <summary>
		/// Prevents a default instance of the <see cref="FileIOHelpers"/> class from being created.
		/// </summary>
		private FileIOHelpers()
		{
		}

		/// <summary>
		/// Serializes the object to XML.
		/// </summary>
		/// <param name="value">The object.</param>
		/// <param name="type">The type.</param>
		/// <returns>System.String.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// value
		/// or
		/// type
		/// </exception>
		public static string SerializeObjectToXml(object value, Type type)
		{
			if (value == null) { throw new ArgumentNullException("value"); }
			if (type == null) { throw new ArgumentNullException("type"); }
			XmlSerializer xser = new XmlSerializer(type);
			return SerializeObjectToXml(value, xser);
		}

		/// <summary>
		/// Serializes the object to XML.
		/// </summary>
		/// <param name="value">The object.</param>
		/// <param name="serializer">The serializer.</param>
		/// <returns>System.String.</returns>
		/// <exception cref="System.ArgumentNullException">
		/// value
		/// or
		/// serializer
		/// </exception>
		public static string SerializeObjectToXml(object value, XmlSerializer serializer)
		{
			if (value == null) { throw new ArgumentNullException("value"); }
			if (serializer == null) { throw new ArgumentNullException("serializer"); }
			StringBuilder sb = new System.Text.StringBuilder(1000);

			using (StringWriter writer = new StringWriter(sb))
			{
				serializer.Serialize(writer, value);
			}

			return sb.ToString();
		}

		/// <summary>
		/// Converts a list of string errors to XML.
		/// </summary>
		/// <param name="errors">The errors.</param>
		/// <param name="rowIndex">Index of the row.</param>
		/// <returns>System.String.</returns>
		public static string ErrorsToXml(IEnumerable<string> errors, int rowIndex)
		{
			var elements = errors.Select(s => new XElement("Error", new XAttribute("row", rowIndex), new XAttribute("message", s)));
			var body = new XElement("Errors", elements);
			return body.ToString();
		}

		/// <summary>
		/// Takes error xml and converts it back to a list of strings.
		/// </summary>
		/// <param name="xml">The XML.</param>
		/// <returns>List&lt;System.String&gt;.</returns>
		public static List<string> ErrorsXmlToList(string xml)
		{
			XDocument doc = XDocument.Parse(xml);

			var list = doc.Root.Elements("Error")
							   .Select(element => element.Attribute("message").Value)
							   .ToList();

			return list;
		}

		/// <summary>
		/// Converts the error xml into a dictionary, with the row number being the key of the dictionary.
		/// </summary>
		/// <param name="xml">The XML.</param>
		/// <returns>Dictionary&lt;System.Int32, List&lt;System.String&gt;&gt;.</returns>
		public static Dictionary<int, List<string>> ErrorsXmlToDictionary(string xml)
		{
			XDocument doc = XDocument.Parse(xml);

			var dict = doc.Root.Elements("Error")
				 .Select(error => new
				 {
					 row = error.Attribute("row").Value,
					 message = error.Attribute("message").Value
				 })
				 .GroupBy(e => e.row)
				 .ToDictionary(g => Convert.ToInt32(g.Key), g =>
				 {
					 return new List<string>(g.Select(err => err.message) as IEnumerable<string>);
				 });

			return dict;
		}

		/// <summary>
		/// Gets the non nullable type of the type.
		/// </summary>
		/// <param name="type">The type.</param>
		/// <returns>Type.</returns>
		public static Type GetNonNullableType(Type type)
		{
			if (Nullable.GetUnderlyingType(type) == null)
			{
				return type;
			}
			else
			{
				return Nullable.GetUnderlyingType(type);
			}
		}

		/// <summary>
		/// Gets the enumerable from reader.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="reader">The reader.</param>
		/// <returns>IEnumerable&lt;T&gt;.</returns>
		/// <exception cref="System.ArgumentNullException">reader</exception>
		public static IEnumerable<T> GetEnumerableFromReader<T>(IDataReader reader)
		{
			if (reader == null) throw new ArgumentNullException("reader");
			StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;

			Type type = typeof(T);

			try
			{
				while (reader.Read())
				{
					T newObj = (T)Activator.CreateInstance(type, true);

					IDataRecordMapper mapper = newObj as IDataRecordMapper;
					if (mapper != null)
					{
						mapper.MapValues(reader);
					}
					else
					{
						var properties = type.GetProperties();
						//resort to reflection to map the object
						for (int i = 0; i < reader.FieldCount; i++)
						{
							if (!reader.IsDBNull(i))
							{
								string name = reader.GetName(i);
								object value = reader.GetValue(i);

								var pi = properties.FirstOrDefault(x => comparer.Equals(x.Name, name));
								if (pi != null)
								{
									pi.SetValue(newObj, value, null);
								}
							}
						}
					}

					yield return newObj;
				}
			}
			finally
			{
				reader.Dispose();
			}
		}


		/// <summary>
		/// Creates a type from the <see cref="IDataReader" /> fields.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <returns>Type.</returns>
		/// <exception cref="System.ArgumentNullException">reader</exception>
		public static Type GetTypeFromReader(IDataRecord record)
		{
			if (record == null) throw new ArgumentNullException("record");

			//https://msdn.microsoft.com/en-us/library/system.reflection.emit.propertybuilder(v=vs.100).aspx
			// create a dynamic assembly and module 
			AssemblyName assemblyName = new AssemblyName("TempReaderAssembly");
			AssemblyBuilder assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
			ModuleBuilder module = assemblyBuilder.DefineDynamicModule("tmpModule");

			// create a new type builder
			TypeBuilder typeBuilder = module.DefineType("ReaderPropertiesType", TypeAttributes.Public | TypeAttributes.Class);

			for (int i = 0; i < record.FieldCount; i++)
			{
				string propertyName = record.GetName(i);
				Type propertyType = record.GetFieldType(i);

				// Generate a private field
				FieldBuilder field = typeBuilder.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);
				// Generate a public property
				PropertyBuilder property = typeBuilder.DefineProperty(propertyName, System.Reflection.PropertyAttributes.HasDefault, propertyType, null);

				// The property set and property get methods require a special set of attributes:
				MethodAttributes methodAttributes = MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig;

				// Define the "get" accessor method for current private field.
				MethodBuilder getterMethod = typeBuilder.DefineMethod("get_" + propertyName, methodAttributes, propertyType, Type.EmptyTypes);
				// Intermediate Language stuff...
				ILGenerator getterIL = getterMethod.GetILGenerator();
				getterIL.Emit(OpCodes.Ldarg_0);
				getterIL.Emit(OpCodes.Ldfld, field);
				getterIL.Emit(OpCodes.Ret);

				// Define the "set" accessor method for current private field.
				MethodBuilder setterMethod = typeBuilder.DefineMethod("set_" + propertyName, methodAttributes, null, new Type[] { propertyType });
				// Again some Intermediate Language stuff...
				ILGenerator setterIL = setterMethod.GetILGenerator();
				setterIL.Emit(OpCodes.Ldarg_0);
				setterIL.Emit(OpCodes.Ldarg_1);
				setterIL.Emit(OpCodes.Stfld, field);
				setterIL.Emit(OpCodes.Ret);

				// Last, we must map the two methods created above to our PropertyBuilder to 
				// their corresponding behaviors, "get" and "set" respectively. 
				property.SetGetMethod(getterMethod);
				property.SetSetMethod(setterMethod);
			}
			Type generatedType = typeBuilder.CreateType();

			return generatedType;

		}

		/// <summary>
		/// Determines whether the test object is a dynamic object.
		/// </summary>
		/// <param name="test">The test.</param>
		/// <returns><c>true</c> if [is dynamic type] [the specified test]; otherwise, <c>false</c>.</returns>
		public static bool IsDynamicType(dynamic test)
		{
			return test is System.Dynamic.IDynamicMetaObjectProvider;
		}
	}

}
