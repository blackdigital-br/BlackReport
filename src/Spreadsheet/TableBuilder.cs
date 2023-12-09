using System.IO;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class TableBuilder
    {
        #region "Constructor"

        internal TableBuilder(WorkbookBuilder spreadsheetBuilder, 
                              SheetBuilder sheetBuilder,
                              ReportConfiguration configuration,
                              string name, 
                              SheetPosition position)
        {
            WorkbookBuilder = spreadsheetBuilder;
            SheetBuilder = sheetBuilder;
            Configuration = configuration;
            TableName = name;
            Position = position;

            WorkbookBuilder.Tables.Add(this);
        }

        #endregion "Constructor"

        #region "Properties"

        private readonly WorkbookBuilder WorkbookBuilder;

        private readonly SheetBuilder SheetBuilder;

        private readonly ReportConfiguration Configuration;

        private readonly string TableName;

        private readonly SheetPosition Position;

        private bool HasData = false;

        private bool HasHeaders = false;

        private ReportSource Header;

        private ReportSource Body;
        
        #endregion "Properties"

        #region "Builder"

        public SheetBuilder Sheet() => SheetBuilder;

        public WorkbookBuilder Workbook() => WorkbookBuilder;

        public Task<ReportFile> BuildAsync() => WorkbookBuilder.BuildAsync();

        public Task BuildAsync(Stream stream) => WorkbookBuilder.BuildAsync(stream);

        public Task BuildAsync(string file) => WorkbookBuilder.BuildAsync(file);

        public TableBuilder AddHeader(ReportSource source)
        {
            if (HasHeaders)
                throw new System.Exception("Table already has headers");

            HasHeaders = true;
            Header = source;
            SheetBuilder.AddValue(source, Position);

            return this;
        }

        public TableBuilder AddHeader<T>(T data)
            => AddHeader(Configuration.Sources.FindSource(data));

        public TableBuilder AddBody(ReportSource source)
        {
            if (HasData)
                throw new System.Exception("Table already has data");

            HasData = true;
            Body = source;
            SheetBuilder.AddValue(source, Position.AddRow());

            return this;
        }

        public TableBuilder AddBody<T>(T data)
            => AddBody(Configuration.Sources.FindSource(data));

        #endregion "Builder"

        #region "Generator"

        public void Generate()
        {
            using MemoryStream tableMemoryStream = new();
            using StreamWriter writer = new(tableMemoryStream);

            //var finalPosition = Position.AddRow(); //Header
            var finalPosition = Position.AddRow(Body.RowCount); //Body
            finalPosition = finalPosition.AddColumn(Body.ColumnCount - 1);

            string tableRef = $"{Position}:{finalPosition}";

            writer.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write($"<x:table xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            writer.Write($"id=\"1\" name=\"{TableName}\" displayName=\"{TableName}\" ");
            writer.Write($"ref=\"{tableRef}\" totalsRowShown=\"0\">");
            writer.Write($"<x:autoFilter ref=\"{tableRef}\" />");
            writer.Write($"<x:tableColumns count=\"{Header.ColumnCount}\">");

            Header.Reset();
            Header.NextRow();

            for (int i = 1; i <= Header.ColumnCount; i++)
            {
                Header.NextColumn();
                writer.Write($"<x:tableColumn id=\"{i}\" name=\"{Header.GetValue()}\" />");
            }

            writer.Write($"</x:tableColumns>");
            writer.Write($"<x:tableStyleInfo name=\"TableStyleMedium15\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" />");
            writer.Write("</x:table>");

            writer.Flush();

            var filename = $"/xl/tables/{WorkbookBuilder.GetTableId(this)}.xml";

            ReportFile file = new(filename,
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                                    tableMemoryStream.ToArray());

            WorkbookBuilder.Files.Add(file);
        }

        #endregion "Generator"
    }
}
