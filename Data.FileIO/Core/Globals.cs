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
using Data.FileIO.Interfaces;
using Microsoft.SqlServer.Server;

namespace Data.FileIO
{
	/// <summary>
	/// Enum ExcelCellType
	/// </summary>
	public enum ExcelCellType
	{
		/// <summary>
		/// The general cell type.
		/// </summary>
		General,
		/// <summary>
		/// The date cell type.
		/// </summary>
		Date,
		/// <summary>
		/// The number cell type.
		/// </summary>
		Number,
		/// <summary>
		/// The decimal cell type.
		/// </summary>
		Decimal,
		/// <summary>
		/// The percent cell type.
		/// </summary>
		Percent,
		/// <summary>
		/// The currency cell type.
		/// </summary>
		Currency,
		/// <summary>
		/// The quoted cell type. Allows for treating a number as text. The format code will be ignored for this option.
		/// </summary>
		Quoted
	}
}