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
        
        //internal Dictionary<uint, Dictionary<uint, SpreadsheetFormatter>> Cells { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public SpreadsheetBuilder Spreadsheet() => SpreadsheetBuilder;

        public Task<byte[]> BuildAsync() => SpreadsheetBuilder.BuildAsync();

        public Task BuildAsync(Stream stream) => SpreadsheetBuilder.BuildAsync(stream);

        public Task BuildAsync(string file) => SpreadsheetBuilder.BuildAsync(file);

        public TableBuilder AddTable(string name, string cellReference)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);

            return AddTable(name, column, row);
        }

        public TableBuilder AddTable(string name, uint column = 1, uint row = 1)
        {
            var table = new TableBuilder(SpreadsheetBuilder, this, name, column, row);
            Tables.Add(table);
            return table;
        }
        
        public SheetBuilder AddValue(object value, string cellReference, string? format = null, IFormatProvider? formatProvider = null)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);

            return AddValue(value, column, row, format, formatProvider);
        }

        public SheetBuilder AddValue(object value, uint column = 1, uint row = 1, string? format = null, IFormatProvider? formatProvider = null)
        {
            return AddValue(value, new SheetPosition(column, row), format, formatProvider);
        }

        public SheetBuilder AddValue(object value, SheetPosition position, string? format = null, IFormatProvider? formatProvider = null)
        {
            SpreadsheetFormatter formatter = new()
            {
                CellReference = SpreadsheetHelper.NumberToExcelColumn(position.Row, position.Column),
                Format = format,
                FormatProvider = formatProvider,
                Value = value,
                ValueType = value?.GetType() ?? typeof(object)
            };

            Values.Add(new SpreadsheetValue(position, new SingleReportValue(value), formatter));

            return this;
        }


        public SheetBuilder FillObject<T>(IEnumerable<T> data, string cellReference, bool generateHeader = true)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);
            return FillObject(data, column, row, generateHeader);
        }

        public SheetBuilder FillObject<T>(IEnumerable<T> data, uint column = 1, uint row = 1, bool generateHeader = true)
        {
            return FillObject(data, new SheetPosition(column, row), generateHeader);
        }

        public SheetBuilder FillObject<T>(IEnumerable<T> data, SheetPosition position, bool generateHeader = true)
        {
            CultureInfo? culture = null;

            if (SpreadsheetBuilder.FormatProvider is CultureInfo cultureInfo)
                culture = cultureInfo;

            return Fill(ReportHelper.ObjectToData(data,
                                                  generateHeader,
                                                  SpreadsheetBuilder.Resource,
                                                  culture), position.Column, position.Row);
        }


        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, string cellReference)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);

            return Fill(data, column, row);
        }

        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, uint column = 1, uint row = 1)
        {
            return Fill(data, new SheetPosition(column, row));
        }

        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, SheetPosition position)
        {
            uint posRow = position.Row;

            foreach (var rowData in data)
            {
                uint posColumn = position.Column;

                foreach (var columnData in rowData)
                {
                    AddValue(columnData, posColumn, posRow);
                    posColumn++;
                }

                posRow++;
            }

            return this;
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
            CreateWorkSheetRowValue(sheetData);

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

        private void CreateWorkSheetRowValue(SheetData sheetData)
        {
            var toProcess = Values.OrderBy(x => x.Position).ToList();

            while (toProcess.Any())
            {
                var row = toProcess.First().Position.Row;
                var processRow = toProcess.Where(x => x.ProcessRow(row)).ToList();

                Row cellRow = new();
                cellRow.RowIndex = (uint)row;

                while (processRow.Any())
                {
                    var column = processRow.First().Position.Column;
                    var processColumn = processRow.Where(x => x.ProcessColumn(column)).ToList(); ;

                    while (processColumn.Any())
                    {
                        var value = processColumn.First();

                        var cellReference = SpreadsheetHelper.NumberToExcelColumn(row, value.Position.Column);
                        cellRow.Append(CreateCellByType(value.Formatter, cellReference));

                        column++;
                        processColumn = processRow.Where(x => x.ProcessColumn(column)).ToList(); ;
                    }

                    row++;
                    processRow = toProcess.Where(x => x.ProcessRow(row)).ToList(); ;
                }

                sheetData.Append(cellRow);
                toProcess.RemoveAll(p => p.Processed);
            }
        }

        
        
        /*private void CreateWorkSheetDataRows(SheetData sheetData)
        {
            foreach (var row in Values.OrderBy(value => value.Position))
            {
                CreateWorksheetRow(row.Key, row.Value, sheetData);
            }
        }

        private static void CreateWorksheetRow(uint row, Dictionary<uint, SpreadsheetFormatter> data, SheetData sheetData)
        {
            Row cellRow = new();
            cellRow.RowIndex = (uint)row;

            foreach (var cellData in data)
            {
                var cellReference = SpreadsheetHelper.NumberToExcelColumn(row, cellData.Key);
                cellRow.Append(CreateCellByType(cellData.Value, cellReference));
            }

            sheetData.Append(cellRow);
        }*/

        private static Cell CreateCellByType(SpreadsheetFormatter value, string cellReference)
        {
            return (new DefaultCellCreate()).CreateCell(value);
        }

        #endregion "Generator"
    }
}
