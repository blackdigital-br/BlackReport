using System.Threading.Tasks;

namespace BlackDigital.Report.Sources
{
    public static class SourceHelper
    {
        public static async Task MoveToEndAsync(this ReportSource source)
        {
            while (await source.NextRowAsync())
            {
                while (await source.NextColumnAsync()) { }
            }
        }
    }
}
