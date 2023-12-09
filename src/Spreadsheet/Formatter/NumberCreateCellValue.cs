using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public class NumberCreateCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = Convert.ToString(Convert.ToDouble(value))
            };
        }
    }

}
