using System.Linq;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public class StringCreaterCellValue : ICreaterCellValue
    {
        public StringCreaterCellValue(SpreadsheetConfiguration configuration)
        {
            Configuration = configuration;
        }

        private SpreadsheetConfiguration Configuration { get; }

        public CellValue Create(SheetPosition position, object? value, HashSet<string>? sharedStrings = null)
        {
            int? ssPosition = null;

            if (sharedStrings != null &&
                sharedStrings.Contains(value?.ToString() ?? string.Empty))
                ssPosition = sharedStrings.ToList().IndexOf(value?.ToString() ?? string.Empty);

            if (ssPosition == null
                && sharedStrings != null
                && value != null
                //&& !value.ToString().Contains('\n')
                && (Configuration.SharedStringsSize.HasValue
                    && Configuration.SharedStringsSize.Value > 0
                    && value.ToString().Length <= Configuration.SharedStringsSize.Value))
            {
                sharedStrings.Add(value.ToString());
                ssPosition = sharedStrings.Count - 1;
            }

            if (ssPosition == null)
            {
                return new()
                {
                    Position = position,
                    Type = CellType.String,
                    Value = value?.ToString() ?? string.Empty
                };
            }

            return new CellValue()
            {
                Position = position,
                Type = CellType.SharedString,
                Value = ssPosition.ToString()
            };
        }
    }
}
