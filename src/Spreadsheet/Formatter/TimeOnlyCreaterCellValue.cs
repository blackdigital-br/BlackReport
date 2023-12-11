using System;
using System.Collections.Generic;
using System.Globalization;

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
            totalSeconds /= 86400;
            string stringValue = Convert.ToString(totalSeconds, CultureInfo.InvariantCulture);

            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = stringValue,
                Style = (int)SpreadsheetFormat.TimeOnly
            };
        }
    }
#endif
}
