// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IFileRowMapper.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data.FileIO.Common.Interfaces
{
	/// <summary>
	/// Used to map file row data to the poco object as the file is read.
	/// </summary>
	public interface IFileRowMapper
	{
		/// <summary>
		/// Maps the values from the dynamic object coming from the file row to the current object.
		/// </summary>
		/// <param name="rowIndex">Index of the row.</param>
		/// <param name="row">The row.</param>
		/// <param name="validator">The validator.</param>
		/// <param name="errors">The errors.</param>
		void MapValues(int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors);
	}
}
