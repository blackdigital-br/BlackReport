using System;
using System.IO;
using System.Resources;
using System.IO.Packaging;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;


namespace BlackDigital.Report.Spreadsheet
{
    public class WorkbookBuilder : ReportBuilder
    {
        #region "Constructor"
        
        public WorkbookBuilder(ReportConfiguration configuration) 
        { 
            Configuration = configuration;
            Creators = new();
            Keywords = new();
        } 

        #endregion "Constructor"

        #region "Properties"

        private string? Application { get; set; }

        private string? Manager { get; set; }

        private string? Company { get; set; }

        private string? AppVersion { get; set; }

        private string? Title { get; set; }

        private string? Subject { get; set; }

        private List<string> Creators { get; set; }

        private List<string> Keywords { get; set; }

        private string? Description { get; set; }

        private string? Category { get; set; }

        protected ReportConfiguration Configuration { get; private set; }

        internal List<SheetBuilder> Sheets { get; private set; } = new();

        internal List<TableBuilder> Tables { get; private set; } = new();

        internal List<ReportFile> Files { get; private set; } = new();

        internal HashSet<string> SharedStrings { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public WorkbookBuilder SetResourceManager(ResourceManager resource) 
            => this.SetResourceManager<WorkbookBuilder>(resource);

        public WorkbookBuilder SetFormatProvider(IFormatProvider formatProvider)
            => this.SetFormatProvider<WorkbookBuilder>(formatProvider);

        public WorkbookBuilder SetApplication(string application)
        {
            Application = application;
            return this;
        }

        public WorkbookBuilder SetManager(string manager)
        {
            Manager = manager;
            return this;
        }

        public WorkbookBuilder SetCompany(string company)
        {
            Company = company;
            return this;
        }

        public WorkbookBuilder SetAppVersion(string appVersion)
        {
            AppVersion = appVersion;
            return this;
        }

        public WorkbookBuilder SetTitle(string title)
        {
            Title = title;
            return this;
        }

        public WorkbookBuilder SetSubject(string subject)
        {
            Subject = subject;
            return this;
        }

        public WorkbookBuilder AddCreators(string creator)
        {
            if (!string.IsNullOrWhiteSpace(creator))
                Creators.Add(creator);

            return this;
        }

        public WorkbookBuilder AddKeyword(string keyword)
        {
            if (!string.IsNullOrWhiteSpace(keyword))
                Keywords.Add(keyword);
            return this;
        }

        public WorkbookBuilder SetDescription(string description)
        {
            Description = description;
            return this;
        }

        public WorkbookBuilder SetCategory(string category)
        {
            Category = category;
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

        public override async Task<ReportFile> BuildAsync()
        {
            using MemoryStream memoryStream = new();
            await CreateSpreadsheetAsync(memoryStream);

            ReportFile file = new($"report_{DateTime.UtcNow}.xlsx",
                    ReportResource.ContentType_Spreadsheet,
                    memoryStream.ToArray());

            return file;
        }

        public override async Task BuildAsync(string file)
        {
            using var fs = File.Create(file);
            await CreateSpreadsheetAsync(fs);
        }

        private async Task CreateSpreadsheetAsync(Stream stream)
        {            
            await GenerateAsync();

            using var package = Package.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            foreach (var file in Files)
            {
                var part = package.CreatePart(new Uri(file.Uri, UriKind.Relative), file.ContentType);
                await part.GetStream().WriteAsync(file.Content, 0, file.Content.Length);
            }

            await stream.FlushAsync();
        }

        private async Task GenerateAsync()
        {
            GenerateDocPropsApp();
            GenerateStyles();

            foreach (var sheet in Sheets)
            {
                await sheet.GenerateAsync();
            }

            foreach (var table in Tables)
            {
                await table.GenerateAsync();
            }

            GenerateSharedStrings();
            GenerateWorkbook();
            GenerateWorkbookRel();
            GeneratePackageRel();
        }

        private void GenerateDocPropsApp()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");

            if (!string.IsNullOrWhiteSpace(Application))
                writer.Write($"<Application>{Application}</Application>");

            if (!string.IsNullOrWhiteSpace(Manager))
                writer.Write($"<Manager>{Manager}</Manager>");

            if (!string.IsNullOrWhiteSpace(AppVersion))
                writer.WriteLine($"<AppVersion>{AppVersion}</AppVersion>");

            writer.Write($"<Company>{Company}</Company>");
            writer.Write("</Properties>");

            writer.Flush();
            var filename = "/docProps/app.xml";

            ReportFile file = new(filename,
                                ReportResource.ContentType_OpenXML_Properties,
                                memoryStream.ToArray());

            Files.Add(file);
        }

        private bool GenerateDocPropsCore()
        {
            if (string.IsNullOrWhiteSpace(Title)
                && string.IsNullOrWhiteSpace(Subject)
                && !Creators.Any()
                && !Keywords.Any()
                && string.IsNullOrWhiteSpace(Description)
                && string.IsNullOrWhiteSpace(Category))
                return false;

            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">");

            if (!string.IsNullOrWhiteSpace(Title))
                writer.Write($"<dc:title>{Title}</dc:title>");

            if (!string.IsNullOrWhiteSpace(Subject))
                writer.Write($"<dc:subject>{Subject}</dc:subject>");

            if (Creators.Any())
                writer.Write($"<dc:creator>{string.Join(";", Creators)}</dc:creator>");

            if (Keywords.Any())
                writer.Write($"<cp:keywords>{string.Join(";", Keywords)}</cp:keywords>");

            if (!string.IsNullOrWhiteSpace(Description))
                writer.Write($"<dc:description>{Description}</dc:description>");

            if (!string.IsNullOrWhiteSpace(Category))
                writer.Write($"<cp:category>{Category}</cp:category>");

            writer.Write("</cp:coreProperties>");

            writer.Flush();
            var filename = "/docProps/core.xml";

            ReportFile file = new(filename,
                                ReportResource.ContentType_OpenXML_Core,
                                memoryStream.ToArray());

            Files.Add(file);
            return true;
        }

        private void GenerateStyles()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write(ReportResource.Spreadsheet_Styles);
            writer.Flush();

            var filename = "/xl/styles.xml";

            ReportFile file = new(filename,
                                  ReportResource.ContentType_Spreadsheet_Styles,
                                  memoryStream.ToArray());

            Files.Add(file);
        }

        private void GenerateSharedStrings()
        {
            if (SharedStrings.Count == 0) return;

            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", SharedStrings.Count);

            foreach (var sharedString in SharedStrings)
            {
                var escapedString = SpreadsheetHelper.Normalize(sharedString);
                writer.Write($"<si><t>{escapedString}</t></si>");
            }

            writer.Write("</sst>");
            writer.Flush();

            var filename = "/xl/sharedStrings.xml";

            ReportFile file = new(filename,
                                  ReportResource.ContentType_Spreadsheet_SharedStrings,
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
                                    ReportResource.ContentType_Spreadsheet_Main,
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

            if (SharedStrings.Count > 0)
            {
                writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"/xl/sharedStrings.xml\" Id=\"sharedStrings1\" />");
            }

            writer.Write("</Relationships>");
            writer.Flush();

            var filename = "/xl/_rels/workbook.xml.rels";

            ReportFile file = new(filename,
                                    ReportResource.ContentType_OpenXML_Relationships,
                                    memoryStream.ToArray());

            Files.Add(file);
        }

        private void GeneratePackageRel()
        {
            using MemoryStream memoryStream = new();
            using StreamWriter writer = new(memoryStream);

            bool containsCore = GenerateDocPropsCore();

            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"/docProps/app.xml\" Id=\"app1\" />");

            if (containsCore)
                writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"/docProps/core.xml\" Id=\"core1\" />");

            writer.Write("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"workbook1\" />");
            writer.Write("</Relationships>");
            writer.Flush();

            var filename = "/_rels/.rels";

            ReportFile file = new(filename,
                                    ReportResource.ContentType_OpenXML_Relationships,
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
