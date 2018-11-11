using Data.DataImport;
using Data.DataImport.Interfaces;
using Data.FileIO.Common.Interfaces;
using Data.FileIO.Common.Utilities;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;

namespace FileIODemo.Data
{
    public partial class Company : ISqlRecordMapper, IFileRowMapper
    {
        #region properties
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        [Range(1, Int32.MaxValue)]
        public int CompanyId { get; set; }

        [Range(typeof(DateTime), "1/1/1753", "12/31/9999")]
        public DateTime StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        [StringLength(150)]
        [Required(AllowEmptyStrings = false)]
        public string LegalName { get; set; }

        [StringLength(150)]
        public string DBAName { get; set; }

        [Range(typeof(DateTime), "1/1/1753", "12/31/9999")]
        public DateTime ChangeDate { get; set; }

        [Required]
        [StringLength(30)]
        public string UserId { get; set; }
        #endregion properties

        #region ISqlRecordMapper Members

        /// <summary>
        /// Provides a custom mapper for the sql record if needed. Typically this might be the case if the columns in the destination table do not match the import type properties.
        /// </summary>
        /// <param name="record">The record.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="errors">The errors.</param>
        public void MapSqlRecord(SqlDataRecord record, int rowIndex, IEnumerable<string> errors)
        {
            record.SetInt32("CompanyId", this.CompanyId);
            Common.SetDateTime(record, "StartDate", this.StartDate);
            Common.SetDateTime(record, "EndDate", this.EndDate);
            record.SetString("LegalName", this.LegalName);
            record.SetString("DBAName", this.DBAName);
            Common.SetDateTime(record, "ChangeDate", this.ChangeDate);
            record.SetString("UserId", this.UserId);

            if (errors.Count() > 0)
            {
                record.SetString("ImportErrors", FileIOHelpers.ErrorsToXml(errors, rowIndex));
            }
        }

        #endregion

        #region IFileRowMapper Members

        /// <summary>
        /// Maps the values.
        /// </summary>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="row">The row.</param>
        /// <param name="validator">The validator.</param>
        /// <param name="errors">The errors.</param>
        public void MapValues(int rowIndex, dynamic row, IObjectValidator validator, ref List<string> errors)
        {
            this.CompanyId = validator.GetRowValue<int>(row, "CompanyId", ref errors);
            this.StartDate = validator.GetRowValue<DateTime>(row, "StartDate", ref errors);
            this.EndDate = validator.GetRowValue<DateTime>(row, "EndDate", ref errors, isNullable: true);
            this.LegalName = row.LegalName;
            this.DBAName = row.DBAName;
            this.ChangeDate = validator.GetRowValue<DateTime>(row, "ChangeDate", ref errors);
            this.UserId = row.UserId;
        }

        #endregion


    }
}
