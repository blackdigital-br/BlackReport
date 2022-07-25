using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public record SpreadsheetFormatter : ValueFormatter
    {
        public HashSet<string> SharedString { get; init; }

        public string CellReference { get; init; }
    }
}
