using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Data.FileIO.Core
{
    internal static class ExcelExtensions
    {
        public static bool IsHidden(this Sheet sheet)
        {
            return sheet.State != null && sheet.State.HasValue &&
                (sheet.State.Value == SheetStateValues.Hidden || sheet.State.Value == SheetStateValues.VeryHidden);
        }
    }
}
