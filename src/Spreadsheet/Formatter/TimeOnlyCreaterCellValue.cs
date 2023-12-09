using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
#if NET6_0_OR_GREATER
    public class TimeOnlyCreaterCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            TimeOnly realValue;

            if (value is TimeOnly timeonly)
                realValue = timeonly;
            else
                throw new InvalidOperationException("Invalid value type");

            double totalSeconds = realValue.Ticks / 10000000;

            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = (totalSeconds / 86400).ToString(),
                Style = (int)SpreadsheetFormat.TimeOnly
            };
        }
    }
#endif
}
