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

        internal Dictionary<uint, Dictionary<uint, object>> Cells { get; private set; } = new();

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
        
        public SheetBuilder AddValue(object value, string cellReference)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);

            return AddValue(value, column, row);
        }

        public SheetBuilder AddValue(object value, uint column = 1, uint row = 1)
        {
            if (!Cells.ContainsKey(row))
                Cells.Add(row, new());

            var rowData = Cells[row];

            if (rowData.ContainsKey(column))
                rowData[column] = value;
            else
                rowData.Add(column, value);

            return this;
        }
        
        public SheetBuilder FillObject<T>(IEnumerable<T> data, string cellReference, bool generateHeader = true)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);
            return FillObject(data, column, row, generateHeader);
        }

        public SheetBuilder FillObject<T>(IEnumerable<T> data, uint column = 1, uint row = 1, bool generateHeader = true)
        {
            CultureInfo? culture = null;

            if (SpreadsheetBuilder.FormatProvider is CultureInfo cultureInfo)
                culture = cultureInfo;

            return Fill(ReportHelper.ObjectToData(data, 
                                                  generateHeader, 
                                                  SpreadsheetBuilder.Resource, 
                                                  culture), column, row);
        }
        
        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, string cellReference)
        {
            var (column, row) = SpreadsheetHelper.CellReferenceToNumbers(cellReference);

            return Fill(data, column, row);
        }

        public SheetBuilder Fill(IEnumerable<IEnumerable<object>> data, uint column = 1, uint row = 1)
        {
            uint posRow = row;

            foreach (var rowData in data)
            {
                uint posColumn = column;

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
            CreateWorkSheetDataRows(sheetData);

            foreach (var table in Tables)
            {
                table.GenerateTablePart(worksheetPart);
            }

            sheets.Append(sheet);
        }

        private void CreateWorksheetColumns(Worksheet worksheet, SheetData sheetData)
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

        private void CreateWorkSheetDataRows(SheetData sheetData)
        {
            foreach (var row in Cells)
            {
                CreateWorksheetRow(row.Key, row.Value, sheetData);
            }
        }

        private void CreateWorksheetRow(uint row, Dictionary<uint, object> data, SheetData sheetData)
        {
            Row cellRow = new();
            cellRow.RowIndex = (uint)row;

            foreach (var cellData in data)
            {
                var letter = SpreadsheetHelper.NumberToExcelColumn(cellData.Key);
                cellRow.Append(CreateCellByType(cellData.Value, $"{letter}{row}"));
            }

            sheetData.Append(cellRow);
        }

        private static Cell CreateCellByType(object value, string cellReference)
        {
            return (new DefaultCellCreate()).CreateCell(new SpreadsheetFormatter
            {
                Value = value,
                ValueType = value?.GetType() ?? typeof(object),
                CellReference = cellReference
            });
        }

        #endregion "Generator"
    }
}
