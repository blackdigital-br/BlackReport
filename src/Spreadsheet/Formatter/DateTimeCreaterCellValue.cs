using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public class DateTimeCreaterCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            DateTime realValue;

            if (value is DateTime dateTime)
                realValue = dateTime;
            else if (value is DateTimeOffset dateTimeOffset)
                realValue = dateTimeOffset.DateTime;
            else
                throw new InvalidOperationException("Invalid value type");

            return new()
            {
                Position = position,
                Type = CellType.Date,
                Value = realValue.ToString("yyyy-MM-ddTHH:mm:ss.fff"),
                Style = (int)SpreadsheetFormat.DateTime
            };
        }
    }
}
