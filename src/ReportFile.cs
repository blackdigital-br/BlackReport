
namespace BlackDigital.Report
{
    public class ReportFile
    {
        public ReportFile(string filename, string contentType, byte[] content)
        {
            Filename = filename;
            ContentType = contentType;
            Content = content;
        }

        public string Filename { get; set; }

        public string ContentType { get; set; }

        public byte[] Content { get; set; }
    }
}
