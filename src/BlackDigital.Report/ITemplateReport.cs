using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public interface ITemplateReport
    {
        Task<byte[]> GenerateReportAsync(string templateName, Dictionary<string, object> data, ReportArguments arguments);

        Task<byte[]> GenerateReportAsync<T>(string templateName, T data, ReportArguments arguments);
    }
}
