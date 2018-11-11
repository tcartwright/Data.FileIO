// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-26-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-26-2016
// ***********************************************************************
// <copyright file="IDataRecordMapper.cs" company="Microsoft">
//     Copyright © Microsoft 2015
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Data.FileIO.Common.Interfaces
{
	/// <summary>
	/// Maps the values in the <see cref="System.Data.IDataRecord" /> to an object.
	/// </summary>
	public interface IDataRecordMapper
	{
		/// <summary>
		/// Maps the values in the <see cref="System.Data.IDataRecord" /> to a single object.
		/// </summary>
		/// <param name="record">The record.</param>
		void MapValues(IDataRecord record);
	}
}
