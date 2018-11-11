// ***********************************************************************
// Assembly         : Data.DataImport
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="ISqlRecordMapper.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SqlServer.Server;

namespace Data.DataImport.Interfaces
{
	/// <summary>
	/// Interface ISqlRecordMapper
	/// </summary>
	public interface ISqlRecordMapper 
	{
		/// <summary>
		/// Provides a custom mapper for the sql record if needed. Typically this might be the case if the columns in the destination table do not match the import type properties.
		/// </summary>
		/// <param name="record">The record.</param>
		/// <param name="rowIndex">Index of the row.</param>
		/// <param name="errors">The errors.</param>
		void MapSqlRecord(SqlDataRecord record, int rowIndex, IEnumerable<string> errors);
	}
}
