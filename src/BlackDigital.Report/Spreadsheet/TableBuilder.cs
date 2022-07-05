using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class TableBuilder
    {
        #region "Constructor"

        internal TableBuilder(SpreadsheetBuilder spreadsheetBuilder, SheetBuilder sheetBuilder, 
            string name, uint column, uint row)
        {
            SpreadsheetBuilder = spreadsheetBuilder;
            SheetBuilder = sheetBuilder;
            TableName = name;
            Row = row;
            Column = column;

            SpreadsheetBuilder.Tables.Add(this);
        }

        #endregion "Constructor"

        #region "Properties"

        private List<string> ColumnsData = new();

        private readonly SpreadsheetBuilder SpreadsheetBuilder;

        private readonly SheetBuilder SheetBuilder;

        private readonly string TableName;

        private readonly uint Row;
        
        private readonly uint Column;

        private bool HasData = false;

        private bool HasHeaders = false;

        private uint Columns = 0;

        private uint Rows = 0;


        #endregion "Properties"

        #region "Builder"

        public SheetBuilder Sheet() => SheetBuilder;

        public SpreadsheetBuilder Spreadsheet() => SpreadsheetBuilder;

        public Task<byte[]> BuildAsync() => SpreadsheetBuilder.BuildAsync();

        public Task BuildAsync(Stream stream) => SpreadsheetBuilder.BuildAsync(stream);

        public Task BuildAsync(string file) => SpreadsheetBuilder.BuildAsync(file);

        public TableBuilder FillObject<T>(IEnumerable<T> data, bool generateHeader = true)
        {
            if (HasData)
                throw new ArgumentException("Table already has data");

            uint row = Row - 1;

            if (generateHeader)
            {
                if (HasHeaders)
                    throw new ArgumentException("Table already has header");

                Rows++;
                HasHeaders = true;
                row--;
            }

            ColumnsData = ReportHelper.GetObjectHeader<T>();
            Columns = (uint)ColumnsData.Count;
            Rows += (uint)data.Count();
            HasData = true;

            SheetBuilder.FillObject(data, Column, Row, generateHeader);

            return this;
        }
        
        public TableBuilder Fill(IEnumerable<IEnumerable<object>> data)
        {
            if (HasData)
                throw new ArgumentException("Table already has data");

            Columns = (uint)data.First().Count();
            Rows += (uint)data.Count();
            HasData = true;

            SheetBuilder.Fill(data, Column, Row + 1);

            return this;
        }

        public TableBuilder AddHeader(IEnumerable<string> headers)
        {
            if (HasHeaders)
                throw new ArgumentException("Table already has header");

            if (Columns < headers.Count())
                Columns = (uint)headers.Count();

            Rows++;
            HasHeaders = true;
            ColumnsData.AddRange(headers);
            List<IEnumerable<object>> data = new();
            data.Add(headers);

            SheetBuilder.Fill(data, Column, Row);

            return this;
        }

        #endregion "Builder"

        #region "Generator"

        internal void GenerateTablePart(WorksheetPart worksheetPart)
        {
            var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();

            Table table = new()
            {
                Id = (uint)SpreadsheetBuilder.Tables.IndexOf(this) + 1,
                Name = TableName,
                DisplayName = TableName,
                TotalsRowShown = false
            };

            var startColumn = SpreadsheetHelper.NumberToExcelColumn(Column);
            var endColumn = SpreadsheetHelper.NumberToExcelColumn(Column + Columns - 1);
            var endRow = Row + Rows - 1;
            var reference = $"{startColumn}{Row}:{endColumn}{endRow}";

            table.Reference = reference;

            AutoFilter autoFilter = new();
            autoFilter.Reference = reference;

            table.Append(autoFilter);

            TableColumns tableColumns = new();
            tableColumns.Count = Columns;

            for (uint i = 0; i < Columns; i++)
            {
                TableColumn tableColumn = new();
                tableColumn.Id = i + 1;
                tableColumn.Name = ColumnsData[(int)i];

                tableColumns.Append(tableColumn);
            }

            table.Append(tableColumns);

            TableStyleInfo tableStyleInfo = new();
            tableStyleInfo.Name = "TableStyleMedium15";
            tableStyleInfo.ShowFirstColumn = false;
            tableStyleInfo.ShowLastColumn = false;
            tableStyleInfo.ShowRowStripes = true;
            tableStyleInfo.ShowColumnStripes = false;

            table.Append(tableStyleInfo);

            tableDefinitionPart.Table = table;

            //TableParts tableParts2 = (TableParts)worksheetPart.Worksheet.ChildElements.Where(ce => ce is TableParts).FirstOrDefault();
            TableParts tableParts = new() { Count = (UInt32)1 };
            TablePart tablePart = new() { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) };

            tableParts.Append(tablePart);

            worksheetPart.Worksheet.Append(tableParts);
        }

        #endregion "Generator"
    }
}
