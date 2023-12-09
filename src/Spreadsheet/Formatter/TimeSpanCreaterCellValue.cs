using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    internal class TimeSpanCreaterCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            TimeSpan realValue;

            if (value is TimeSpan timeSpan)
                realValue = timeSpan;
            else
                throw new InvalidOperationException("Invalid value type");

            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = (realValue.TotalSeconds / 86400).ToString(),
                Style = (int)SpreadsheetFormat.TimeSpan
            };
        }
    }
}
