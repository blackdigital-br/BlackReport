using System.Collections.Generic;


namespace BlackDigital.Report.Spreadsheet
{
    public record SpreadsheetFormatter : ValueFormatter
    {
        public HashSet<string>? SharedString { get; init; }

        public SpreadsheetFormat? FormatType { get; init; }
    }
}
