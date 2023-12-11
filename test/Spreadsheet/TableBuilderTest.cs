using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Sources;
using BlackDigital.Report.Spreadsheet;

namespace BlackDigital.Report.Tests.Spreadsheet
{
    public class TableBuilderTest
    {
        [Fact]
        public async void GenerateTable()
        {
            var spreadsheet = new WorkbookBuilder(new ReportConfiguration());
            var sheet = spreadsheet.AddSheet("Sheet1");
            var table = sheet.AddTable("Table1", 1, 1);

            var header = new ListReportSource(new string[]
            {
                "Header1",
                "Header2"
            });

            var body = new ModelReportSource<SimpleModel>(new List<SimpleModel>
            {
                new("Hello", 1),
                new("World", 2)
            });

            table.AddHeader(header);
            table.AddBody(body);

            await header.MoveToEndAsync();
            await body.MoveToEndAsync();

            await table.GenerateAsync();

            Assert.Contains(spreadsheet.Files, f => f.Filename == "/xl/tables/table1.xml");

            using MemoryStream msTable = new(spreadsheet.Files.First(f => f.Filename == "/xl/tables/table1.xml").Content);
            using StreamReader srTable = new(msTable);
            var tableXml = srTable.ReadToEnd();

            var expected = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:B3\" totalsRowShown=\"0\"><autoFilter ref=\"A1:B3\" /><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Header1\" /><tableColumn id=\"2\" name=\"Header2\" /></tableColumns><tableStyleInfo name=\"TableStyleMedium15\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /></table>";
            Assert.Equal(expected, tableXml);
        }
    }
}
