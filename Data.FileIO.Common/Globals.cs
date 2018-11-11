// ***********************************************************************
// Assembly         : Data.FileIO
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="Globals.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data.FileIO.Common.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.FileIO.Common
{
	/// <summary>
	/// Used to provide custom mapping when reading files and mapping the contents of the file to the poco.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="mapObject">The map object.</param>
	/// <param name="rowIndex">Index of the row.</param>
	/// <param name="row">The row.</param>
	/// <param name="validator">The validator.</param>
	/// <param name="errors">The errors.</param>
	public delegate void FileValuesMap<T>(ref T mapObject, int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors);
}