
using System;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public enum CellType
    {
        Boolean,
        Number,
        Error,
        SharedString,
        String,
        Date
    }

    public static class CellTypeExtensions
    {
        public static string ToCellTypeString(this CellType cellType)
        {
            return cellType switch
            {
                CellType.Boolean => "b",
                CellType.Number => "n",
                CellType.Error => "e",
                CellType.SharedString => "s",
                CellType.String => "str",
                CellType.Date => "d",
                _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
            };
        }
    }
}
