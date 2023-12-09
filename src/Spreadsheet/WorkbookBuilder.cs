using System;
using System.IO;
using System.Resources;
using System.IO.Packaging;
using System.Threading.Tasks;
using System.Collections.Generic;


namespace BlackDigital.Report.Spreadsheet
{
    public class WorkbookBuilder : ReportBuilder
    {
        #region "Constructor"
        
        public WorkbookBuilder(ReportConfiguration configuration) 
        { 
            Configuration = configuration;
        } 

        #endregion "Constructor"

        #region "Properties"

        private string Company { get; set; } = null;

        protected ReportConfiguration Configuration { get; private set; }

        internal List<SheetBuilder> Sheets { get; private set; } = new();

        internal List<TableBuilder> Tables { get; private set; } = new();

        internal List<ReportFile> Files { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public WorkbookBuilder SetResourceManager(ResourceManager resource) 
            => this.SetResourceManager<WorkbookBuilder>(resource);

        public WorkbookBuilder SetFormatProvider(IFormatProvider formatProvider)
            => this.SetFormatProvider<WorkbookBuilder>(formatProvider);

        public WorkbookBuilder SetCompany(string company)
        {
            Company = company;
            return this;
        }

        public SheetBuilder AddSheet(string name)
        {
            var sheet = new SheetBuilder(this, Configuration, name);
            Sheets.Add(sheet);
            return sheet;
        }

        #endregion "Builder"

        #region "Generate"

        public override Task<ReportFile> BuildAsync()
        {
            return Task.Run(() => {
                using MemoryStream memoryStream = new();
                CreateSpreadsheet(memoryStream);

                ReportFile file = new($"report_{DateTime.UtcNow}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        memoryStream.ToArray());

                return file;
            });
        }

        public override Task BuildAsync(string file)
        {
            return Task.Run(() => {
                using var fs = File.Create(file);
                CreateSpreadsheet(fs);
            });
        }

        private void CreateSpreadsheet(Stream stream)
        {
            /*using var package = Package.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using SpreadsheetDocument document = SpreadsheetDocument.Create(package, DocumentType);*/
            //CreateParts(document);
            
            Generate();

            using var package = Package.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            foreach (var file in Files)
            {
                var part = package.CreatePart(new Uri(file.Uri, UriKind.Relative), file.ContentType);
                part.GetStream().Write(file.Content, 0, file.Content.Length);
            }

            stream.Flush();
        }

        private void Generate()
        {
            GenerateDocProps();
            GenerateStyles();

            foreach (var sheet in Sheets)
            {
                sheet.Generate();
            }

            foreach (var table in Tables)
            {
                table.Generate();
            }

            GenerateWorkbook();
            GenerateWorkbookRel();
            GeneratePackageRel();
        }

        private void GenerateDocProps()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<ap:Properties xmlns:ap=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
            writer.Write($"<ap:Company>{Company}</ap:Company>");
            writer.Write("</ap:Properties>");

            writer.Flush();
            var filename = "/docProps/app.xml";

            ReportFile file = new(filename,
                                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                                memoryStream.ToArray());

            Files.Add(file);
        }

        private void GenerateStyles()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write(SpreadsheetResource.Styles);
            writer.Flush();

            var filename = "/xl/styles.xml";

            ReportFile file = new(filename,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                                  memoryStream.ToArray());

            Files.Add(file);
        }

        private void GenerateWorkbook()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write("<x:sheets>");

            for(int i = 1; i <= Sheets.Count; i++)
            {
                var sheet = Sheets[i - 1];
                var sheetId = Sheets.IndexOf(sheet) + 1;
                writer.Write($"<x:sheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" name=\"{sheet.SheetName}\" sheetId=\"{sheetId}\" r:id=\"{GetSheetId(sheet)}\" />");
            }

            writer.Write("</x:sheets>");
            writer.Write("</x:workbook>");
            
            writer.Flush();

            var filename = "/xl/workbook.xml";

            ReportFile file = new(filename,
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                                    memoryStream.ToArray());

            Files.Add(file);
        }

        private void GenerateWorkbookRel()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"style1\" />");

            for (int i = 1; i <= Sheets.Count; i++)
            {
                var sheet = Sheets[i - 1];
                var sheetId = GetSheetId(sheet);
                writer.Write($"<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/{sheetId}.xml\" Id=\"{sheetId}\" />");
            }

            writer.Write("</Relationships>");
            writer.Flush();

            var filename = "/xl/_rels/workbook.xml.rels";

            ReportFile file = new(filename,
                                    "application/vnd.openxmlformats-package.relationships+xml",
                                    memoryStream.ToArray());

            Files.Add(file);
        }

        private void GeneratePackageRel()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"/docProps/app.xml\" Id=\"app1\" />");
            writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"workbook1\" />");
            writer.Write("</Relationships>");
            writer.Flush();

            var filename = "/_rels/.rels";

            ReportFile file = new(filename,
                                    "application/vnd.openxmlformats-package.relationships+xml",
                                    memoryStream.ToArray());

            Files.Add(file);
        }

        internal string GetTableId(TableBuilder table)
        {
            return $"table{Tables.IndexOf(table) + 1}";
        }

        internal string GetSheetId(SheetBuilder sheet)
        {
            return $"sheet{Sheets.IndexOf(sheet) + 1}";
        }

        #endregion "Generate"
    }
}
