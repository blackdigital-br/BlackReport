using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public class BooleanCreaterCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            return new()
            {
                Position = position,
                Type = CellType.Boolean,
                Value = Convert.ToBoolean(value) ? "1" : "0"
            };
        }
    }
}
