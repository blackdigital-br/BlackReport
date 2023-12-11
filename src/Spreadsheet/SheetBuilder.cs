using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using BlackDigital.Report.Spreadsheet.Formatter;

namespace BlackDigital.Report.Spreadsheet
{
    public class SheetBuilder
    {
        #region "Constructor"

        internal SheetBuilder(WorkbookBuilder spreadsheetBuilder, 
                            ReportConfiguration configuration,
                            string name)
        {
            SheetName = name;
            WorkbookBuilder = spreadsheetBuilder;
            Configuration = configuration;
            Values = new();
        }

        #endregion "Constructor"

        #region "Properties"

        private readonly WorkbookBuilder WorkbookBuilder;

        private readonly ReportConfiguration Configuration;

        internal protected readonly string SheetName;

        protected List<TableBuilder> Tables { get; private set; } = new();

        protected Dictionary<SheetPosition, ReportSource> Values;

        //internal List<SpreadsheetValue> Values { get; private set; } = new();
        
        //internal Dictionary<uint, Dictionary<uint, ValueFormatter>> Cells { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public WorkbookBuilder Workbook() => WorkbookBuilder;

        public Task<ReportFile> BuildAsync() => WorkbookBuilder.BuildAsync();

        public Task BuildAsync(Stream stream) => WorkbookBuilder.BuildAsync(stream);

        public Task BuildAsync(string file) => WorkbookBuilder.BuildAsync(file);


        public TableBuilder AddTable(string name, string cellReference)
        {
            return AddTable(name, (SheetPosition)cellReference);
        }

        public TableBuilder AddTable(string name, uint column = 1, uint row = 1)
        {
            return AddTable(name, new SheetPosition(column, row));
        }

        public TableBuilder AddTable(string name, SheetPosition position)
        {
            var table = new TableBuilder(WorkbookBuilder, 
                                         this,
                                         Configuration,
                                         name, 
                                         position);
            Tables.Add(table);
            return table;
        }

        public SheetBuilder AddValue(ReportSource source, string cellReference)
            => AddValue(source, (SheetPosition)cellReference);
        
        public SheetBuilder AddValue(ReportSource source, uint column = 1, uint row = 1)
            => AddValue(source, new SheetPosition(column, row));

        public SheetBuilder AddValue(ReportSource source, SheetPosition position)
        {
            if (Values.ContainsKey(position))
                Values[position] = source;
            else
                Values.Add(position, source);

            return this;
        }

        public SheetBuilder AddValue<T>(T data, uint column = 1, uint row = 1)
            => AddValue<T>(data, new SheetPosition(column, row));

        public SheetBuilder AddValue<T>(T data, string cellReference)
            => AddValue<T>(data, (SheetPosition)cellReference);

        public SheetBuilder AddValue<T>(T value, SheetPosition position)
            => AddValue(Configuration.Sources.FindSource<T>(value), position);

        #endregion "Builder"

        #region "Generator"

        /// <summary>
        /// Generate sheets files
        /// </summary>
        /// <param name="configuration"></param>
        public async Task GenerateAsync()
        {
            await GenerateSheetAsync();
            GenerateSheetRels();
        }

        /// <summary>
        /// Create a sheet.xml file
        /// </summary>
        /// <param name="configuration"></param>
        private async Task GenerateSheetAsync()
        {
            using MemoryStream sheetMemoryStream = new();
            using StreamWriter writer = new(sheetMemoryStream);

            writer.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write($"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write($"<sheetData>");
            await WriteSheetAsync(writer);
            writer.Write($"</sheetData>");
            WriteTableParts(writer);
            writer.Write($"</worksheet>");

            writer.Flush();

            var filename = $"/xl/worksheets/{WorkbookBuilder.GetSheetId(this)}.xml";

            ReportFile file = new(filename,
                                    ReportResource.ContentType_Spreadsheet_Sheet,
                                    sheetMemoryStream.ToArray());

            WorkbookBuilder.Files.Add(file);
        }

        /// <summary>
        /// Create a sheet.xml.rels file
        /// </summary>
        private void GenerateSheetRels()
        {
            using MemoryStream sheetMemoryStream = new();
            using StreamWriter writer = new(sheetMemoryStream);

            writer.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.Write($"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");

            foreach (var table in Tables)
            {
                string tableId = WorkbookBuilder.GetTableId(table);
                writer.Write($"<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"/xl/tables/{tableId}.xml\" Id=\"{tableId}\" />");
            }

            writer.Write($"</Relationships>");

            writer.Flush();

            var filename = $"/xl/worksheets/_rels/{WorkbookBuilder.GetSheetId(this)}.xml.rels";

            ReportFile file = new(filename,
                                    ReportResource.ContentType_OpenXML_Relationships,
                                    sheetMemoryStream.ToArray());

            WorkbookBuilder.Files.Add(file);
        }

        /// <summary>
        /// Find sources to write
        /// </summary>
        /// <param name="writer"></param>
        private async Task WriteSheetAsync(StreamWriter writer)
        {
            var sourceOrdened = Values.OrderBy(source => source.Key.Row)
                                      .ThenBy(x => x.Key.Column)
                                      .ToDictionary(x => x.Key, x => x.Value);

            SheetPosition position = sourceOrdened.First().Key;

            while (sourceOrdened.Any())
            {
                Dictionary<SheetPosition, ReportSource> sources = new();

                foreach (var source in sourceOrdened.ToArray())
                {
                    if (position.Row >= source.Key.Row)
                    {
                        if (await source.Value.NextRowAsync())
                            sources.Add(source.Key, source.Value);
                        else
                            sourceOrdened.Remove(source.Key);
                    }
                }

                position = await WriteRowAsync(writer, sources, position);
                position = new(1, position.Row + 1);
            }
        }

        /// <summary>
        /// Write a row to sheet.xml file
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="sources"></param>
        /// <param name="position"></param>
        private async Task<SheetPosition> WriteRowAsync(StreamWriter writer, 
                              Dictionary<SheetPosition, ReportSource> sources,
                              SheetPosition position)
        {
            if (!sources.Any())
                return position;

            writer.Write($"<row r=\"{position.Row}\">");

            while (sources.Any())
            {
                Dictionary<SheetPosition, ReportSource> columnSources = new();

                foreach (var source in sources.ToArray())
                {
                    if (position.Column >= source.Key.Column)
                    {
                        if (await source.Value.NextColumnAsync())
                            columnSources.Add(source.Key, source.Value);
                        else
                            sources.Remove(source.Key);
                    }
                }

                position = await WriteColumnAsync(writer, columnSources, position);
                position = position.AddColumn();
            }
            

            writer.Write("</row>");
            writer.Flush();

            return position;
        }

        /// <summary>
        /// Write a cell in column to sheet.xml file
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="sources"></param>
        /// <param name="position"></param>
        private async Task<SheetPosition> WriteColumnAsync(StreamWriter writer,
                                 Dictionary<SheetPosition, ReportSource> sources,
                                 SheetPosition position)
        {
            if (!sources.Any())
                return position;
                
            var value = await sources.First().Value.GetValueAsync();
            var createCellValue = Configuration.Spreadsheet.GetCreaterCellValue(value?.GetType());

            if (createCellValue is null)
                return position;

            var valueCell = createCellValue.Create(position, value, WorkbookBuilder.SharedStrings);

            string style = valueCell.Style.HasValue ? $" s=\"{valueCell.Style.Value}\"" : string.Empty;
            string finalValue = SpreadsheetHelper.Normalize(valueCell.Value.ToString());

            writer.Write($"<c r=\"{valueCell.Position}\"{style} t=\"{valueCell.Type.ToCellTypeString()}\">");
            writer.Write($"<v>{finalValue}</v></c>");

            return position;
        }

        /// <summary>
        /// Write tableParts to sheet.xml file
        /// </summary>
        /// <param name="writer"></param>
        private void WriteTableParts(StreamWriter writer)
        {
            if (Tables.Any())
            {
                writer.Write($"<tableParts count=\"{Tables.Count}\">");

                foreach (var table in Tables)
                {
                    writer.Write($"<tablePart xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{WorkbookBuilder.GetTableId(table)}\" />");
                }

                writer.Write("</tableParts>");
            }   
        }

        #endregion "Generator"
    }
}
