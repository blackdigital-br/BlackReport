using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class SheetBuilder
    {
        #region "Constructor"

        internal SheetBuilder(SpreadsheetBuilder spreadsheetBuilder, string name)
        {
            SheetName = name;
            SpreadsheetBuilder = spreadsheetBuilder;
        }

        #endregion "Constructor"

        #region "Properties"

        private readonly SpreadsheetBuilder SpreadsheetBuilder;

        private readonly string SheetName;

        internal List<TableBuilder> Tables { get; private set; } = new();

        internal List<SpreadsheetValue> Values { get; private set; } = new();
        
        //internal Dictionary<uint, Dictionary<uint, ValueFormatter>> Cells { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public SpreadsheetBuilder Spreadsheet() => SpreadsheetBuilder;

        public Task<byte[]> BuildAsync() => SpreadsheetBuilder.BuildAsync();

        public Task BuildAsync(Stream stream) => SpreadsheetBuilder.BuildAsync(stream);

        public Task BuildAsync(string file) => SpreadsheetBuilder.BuildAsync(file);

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
            var table = new TableBuilder(SpreadsheetBuilder, this, name, position);
            Tables.Add(table);
            return table;
        }

        public SheetBuilder AddValue(object value, string cellReference, string? format = null, IFormatProvider? formatProvider = null)
        {
            return AddValue(value, new SheetPosition(cellReference), format, formatProvider);
        }

        public SheetBuilder AddValue(object value, uint column = 1, uint row = 1, string? format = null, IFormatProvider? formatProvider = null)
        {
            return AddValue(value, new SheetPosition(column, row), format, formatProvider);
        }

        public SheetBuilder AddValue(object value, SheetPosition position, string? format = null, IFormatProvider? formatProvider = null)
        {
            ValueFormatter formatter = new()
            {
                Format = format,
                FormatProvider = formatProvider
            };

            Values.Add(new SpreadsheetValue(position, new SingleReportSource(value), formatter));

            return this;
        }


        public SheetBuilder FillObject<T>(IEnumerable<T> data, string cellReference, bool generateHeader = true)
        {
            return FillObject(data, new SheetPosition(cellReference), generateHeader);
        }

        public SheetBuilder FillObject<T>(IEnumerable<T> data, uint column = 1, uint row = 1, bool generateHeader = true)
        {
            return FillObject(data, new SheetPosition(column, row), generateHeader);
        }

        public SheetBuilder FillObject<T>(IEnumerable<T> data, SheetPosition position, bool generateHeader = true)
        {
            /*CultureInfo? culture = null;

            if (SpreadsheetBuilder.FormatProvider is CultureInfo cultureInfo)
                culture = cultureInfo;*/

            InternalFillObject(data, position, generateHeader);
            return this;
            /*return Fill(ReportHelper.ObjectToData(data,
                                                  generateHeader,
                                                  SpreadsheetBuilder.Resource,
                                                  culture), position.Column, position.Row);*/
        }

        internal SpreadsheetValue InternalFillObject<T>(IEnumerable<T> data, SheetPosition position, bool generateHeader)
        {
            var bodyPosition = position;

            if (generateHeader)
            {
                var headerValue = new SpreadsheetValue(position, ReportHelper.GetObjectHeader<T>(null, null));
                Values.Add(headerValue);
                bodyPosition.AddRow(1);
            }

            var value = new SpreadsheetValue(bodyPosition, new ModelReportSource<T>(data), null);
            Values.Add(value);
            return value;
        }


        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, string cellReference)
        {
            return Fill(data, new SheetPosition(cellReference));
        }

        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, uint column = 1, uint row = 1)
        {
            return Fill(data, new SheetPosition(column, row));
        }

        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, SheetPosition position)
        {
            Values.Add(new SpreadsheetValue(position, new EnumerableReportSource(data)));
            return this;
        }

        internal SpreadsheetValue InternalFill(IEnumerable<IEnumerable<object>> data, SheetPosition position)
        {
            var value = new SpreadsheetValue(position, new EnumerableReportSource(data));
            Values.Add(value);
            return value;
        }

        #endregion "Builder"

        #region "Generator"

        internal void GenerateWorksheetPart(WorkbookPart workbookPart, Sheets sheets, HashSet<string> sharedStrings)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new();
            Worksheet worksheet = new(sheetData);
            worksheetPart.Worksheet = worksheet;

            Sheet sheet = new()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)(SpreadsheetBuilder.Sheets.IndexOf(this) + 1),
                Name = SheetName
            };

            CreateWorksheetColumns(worksheet, sheetData);
            CreateWorkSheetValues(sheetData);

            foreach (var table in Tables)
            {
                table.GenerateTablePart(worksheetPart);
            }

            sheets.Append(sheet);
        }

        private static void CreateWorksheetColumns(Worksheet worksheet, SheetData sheetData)
        {
            /*Columns xColumns = new();

            foreach (string column in columns)
            {
                Column xColumn = new();
                xColumn.Width = 15;
                xColumn.Min = 1;
                xColumn.Max = 1;
                xColumn.CustomWidth = true;
                xColumns.Append(xColumn);
            }

            xWorksheet.Append(xColumns);*/
            //CreateWorksheetRow(columns, xSheetData, 1);
        }

        private void CreateWorkSheetValues(SheetData sheetData)
        {
            var toProcess = Values.OrderBy(x => x.Position).ToList();

            while (toProcess.Any())
            {
                var row = toProcess.First().Position.Row;
                var processRow = toProcess.Where(x => x.ProcessRow(row)).ToList();

                while (processRow.Any())
                {
                    Row cellRow = new();
                    cellRow.RowIndex = (uint)row;

                    var column = processRow.First().Position.Column;
                    var processColumn = processRow.Where(x => x.ProcessColumn(column)).ToList();

                    while (processColumn.Any())
                    {
                        var value = processColumn.First();
                        cellRow.Append(CreateCellByType(new SheetPosition(column, row), value.Value.GetValue(), value.GetFormatter(row, column)));

                        column++;
                        processColumn = processRow.Where(x => x.ProcessColumn(column)).ToList();
                    }

                    sheetData.Append(cellRow);

                    row++;
                    processRow = toProcess.Where(x => x.ProcessRow(row)).ToList();
                }

                toProcess.RemoveAll(p => p.Processed);
            }
        }

        private static Cell CreateCellByType(SheetPosition position, object? value, ValueFormatter formatter)
        {
            return (new DefaultCellCreate()).CreateCell(position, value, formatter);
        }

        #endregion "Generator"
    }
}
