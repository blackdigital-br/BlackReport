
namespace BlackDigital.Report
{
    public class ReportFile
    {
        public ReportFile(string uri, string contentType, byte[] content)
        {
            Uri = uri;
            ContentType = contentType;
            Content = content;
        }

        public string Uri { get; set; }

        public string ContentType { get; set; }

        public byte[] Content { get; set; }
    }
}
