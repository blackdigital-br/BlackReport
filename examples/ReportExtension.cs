using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Example
{
    public static class ReportExtension
    {
        public static WorkbookBuilder MyReport<TModel>(this WorkbookBuilder builder, string name, IEnumerable<TModel> list)
        {
            return null;
            /*return builder.SetCompany("My Company")
                          .AddSheet(name)
                          .AddValue("My Company")
                          .AddValue("Date: ", "A2")
                          .AddValue(DateTime.Now, "B2")
                          .AddValue(name, "A3")
                          .AddTable("report", "A5")
                          .FillObject(list)
                          .Spreadsheet();*/
        }

        public static async Task<string> BuilderReportAsync(this WorkbookBuilder builder)
        {
            string filename = Guid.NewGuid().ToString();
            filename = filename.Replace("-", "");
            filename = $"{filename}.xlsx";

            await builder.BuildAsync($"/var/www/downloads/reports/{filename}");

            return filename;
        }
    }
}
