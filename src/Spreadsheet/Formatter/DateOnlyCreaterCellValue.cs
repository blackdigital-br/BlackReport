using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
#if NET6_0_OR_GREATER
    public class DateOnlyCreaterCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            DateOnly realValue;

            if (value is DateOnly dateonly)
                realValue = dateonly;
            else
                throw new InvalidOperationException("Invalid value type");

            return new()
            {
                Position = position,
                Type = CellType.Date,
                Value = realValue.ToDateTime(TimeOnly.MinValue).ToString("yyyy-MM-ddT00:00:00.000"),
                Style = (int)SpreadsheetFormat.DateOnly
            };
        }
    }
#endif
}
