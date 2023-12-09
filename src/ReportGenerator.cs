using BlackDigital.Report.Spreadsheet;


namespace BlackDigital.Report
{
    public static class ReportGenerator
    {
        public static WorkbookBuilder Spreadsheet() => new(new ReportConfiguration());

        public static WorkbookBuilder Spreadsheet(this Report report) 
            => new(report.Configuration);
    }
}
