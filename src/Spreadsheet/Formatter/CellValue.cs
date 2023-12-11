
namespace BlackDigital.Report.Spreadsheet.Formatter
{
    public struct CellValue
    {
        public SheetPosition Position { get; set; }

        public int? Style { get; set; }

        public CellType Type { get; set; }

        public string Value { get; set; }
    }
}
