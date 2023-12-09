
namespace BlackDigital.Report.Spreadsheet
{
    public enum SpreadsheetFormat : uint
    {
        None = 0,
        DateOnly = 1,
        TimeSpan = 2,
        DateTime = 3,
        TimeOnly = 4,
        Custom = uint.MaxValue
    }
}
