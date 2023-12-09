
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public interface ICreaterCellValue
    {
        CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null);
    }
}
