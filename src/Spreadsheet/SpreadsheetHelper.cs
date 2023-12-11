
namespace BlackDigital.Report.Spreadsheet
{
    internal static class SpreadsheetHelper
    {
        internal static string Normalize(string xml)
        {
            return xml.Replace("\r\n", "\n")
                      .Replace("&", "&amp;")
                      .Replace("<", "&lt;")
                      .Replace(">", "&gt;");
        }
    }
}
