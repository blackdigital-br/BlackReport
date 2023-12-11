using System;
using System.Collections.Generic;
using System.Globalization;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public class NumberCreateCellValue : ICreaterCellValue
    {
        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            double realValue = Convert.ToDouble(value);
            string stringValue = Convert.ToString(realValue, CultureInfo.InvariantCulture);

            return new()
            {
                Position = position,
                Type = CellType.Number,
                Value = stringValue
            };
        }
    }

}
