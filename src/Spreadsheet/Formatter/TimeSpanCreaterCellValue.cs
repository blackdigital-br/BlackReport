using System;
using System.Collections.Generic;
using System.Globalization;

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

            double totalSeconds = realValue.TotalSeconds / 86400;
            string stringValue = Convert.ToString(totalSeconds, CultureInfo.InvariantCulture);

            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = stringValue,
                Style = (int)SpreadsheetFormat.TimeSpan
            };
        }
    }
}
