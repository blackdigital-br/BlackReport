using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Sources;
using BlackDigital.Report.Spreadsheet;

namespace BlackDigital.Report.Tests.Spreadsheet
{
    public class SheetBuilderTest
    {
        [Fact]
        public void GenerateSheet()
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

            sheet.Generate();

            Assert.Contains("xl/worksheets/sheet1.xml", spreadsheet.Files.Keys);
            Assert.Contains("xl/worksheets/_rels/sheet1.xml.rels", spreadsheet.Files.Keys);

            using MemoryStream msSheet = new(spreadsheet.Files["xl/worksheets/sheet1.xml"]);
            using StreamReader srSheet = new(msSheet);
            var sheetXml = srSheet.ReadToEnd();

            var expected = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\" t=\"str\">Hello</x:c><x:c r=\"B1\" t=\"str\">1</x:c></x:row><x:row r=\"2\"><x:c r=\"A2\" t=\"str\">World</x:c><x:c r=\"B2\" t=\"str\">2</x:c></x:row></x:sheetData></x:worksheet>";
            Assert.Equal(expected, sheetXml);

            using MemoryStream msSheetRels = new(spreadsheet.Files["xl/worksheets/_rels/sheet1.xml.rels"]);
            using StreamReader srSheetRels = new(msSheetRels);

            var sheetRelsXml = srSheetRels.ReadToEnd();
            expected = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";

            Assert.Equal(expected, sheetRelsXml);
        }
    }
}
