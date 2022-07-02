namespace BlackDigital.Report
{
    public interface IReport
    {
        Task<byte[]> GenerateReportAsync(IEnumerable<IEnumerable<object>> data, IEnumerable<string> columns, ReportArguments arguments);
        
        Task<byte[]> GenerateReportAsync<T>(IEnumerable<T> data, ReportArguments arguments);        
    }
}