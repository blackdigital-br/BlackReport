using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Sources;
using BlackDigital.Report.Spreadsheet;

namespace BlackDigital.Report.Tests.Spreadsheet
{
    public class SheetBuilderTest
    {
        [Fact]
        public async void GenerateSheet()
        {
            var spreadsheet = new WorkbookBuilder(new ReportConfiguration());
            var sheet = spreadsheet.AddSheet("Sheet1");
            var value = new ModelReportSource<SimpleModel>();

            value.Load(new List<SimpleModel>
            {
                new("Hello", 1),
                new("World", 2)
            });

            sheet.AddValue(value, new SheetPosition(1, 1));

            await sheet.GenerateAsync();

            Assert.Contains(spreadsheet.Files, f => f?.Filename == "/xl/worksheets/sheet1.xml");
            Assert.Contains(spreadsheet.Files, f => f?.Filename == "/xl/worksheets/_rels/sheet1.xml.rels");

            using MemoryStream msSheet = new(spreadsheet.Files.First(f => f.Filename == "/xl/worksheets/sheet1.xml").Content);
            using StreamReader srSheet = new(msSheet);
            var sheetXml = srSheet.ReadToEnd();

            var expected = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"n\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c><c r=\"B2\" t=\"n\"><v>2</v></c></row></sheetData></worksheet>";
            Assert.Equal(expected, sheetXml);

            using MemoryStream msSheetRels = new(spreadsheet.Files.First(f => f.Filename == "/xl/worksheets/_rels/sheet1.xml.rels").Content);
            using StreamReader srSheetRels = new(msSheetRels);

            var sheetRelsXml = srSheetRels.ReadToEnd();
            expected = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";

            Assert.Equal(expected, sheetRelsXml);
        }
    }
}
