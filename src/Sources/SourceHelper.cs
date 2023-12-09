
namespace BlackDigital.Report.Sources
{
    public static class SourceHelper
    {
        public static void MoveToEnd(this ReportSource source)
        {
            while (source.NextRow())
            {
                while (source.NextColumn()) { }
            }
        }
    }
}
