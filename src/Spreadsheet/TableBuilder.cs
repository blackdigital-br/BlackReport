using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class TableBuilder
    {
        #region "Constructor"

        internal TableBuilder(SpreadsheetBuilder spreadsheetBuilder, SheetBuilder sheetBuilder, 
            string name, SheetPosition position)
        {
            SpreadsheetBuilder = spreadsheetBuilder;
            SheetBuilder = sheetBuilder;
            TableName = name;
            Position = position;

            SpreadsheetBuilder.Tables.Add(this);
        }

        #endregion "Constructor"

        #region "Properties"

        private readonly SpreadsheetBuilder SpreadsheetBuilder;

        private readonly SheetBuilder SheetBuilder;

        private readonly string TableName;

        private readonly SheetPosition Position;

        private bool HasData = false;

        private bool HasHeaders = false;

        private SpreadsheetValue HeaderValue;

        private SpreadsheetValue BodyValue;
        
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

            if (generateHeader)
            {
                if (HasHeaders)
                    throw new ArgumentException("Table already has header");

                HeaderValue = new SpreadsheetValue(Position, ReportHelper.GetObjectHeader<T>(null, null));
                HasHeaders = true;
            }

            if (!HasHeaders)
                throw new ArgumentException("Table has no header");

            /*CultureInfo? culture = null;
            
            if (SpreadsheetBuilder.FormatProvider is CultureInfo cultureInfo)
                culture = cultureInfo;*/

            var position = Position;

            if (HasHeaders)
                position = position.AddRow(1);

            HasData = true;
            BodyValue = SheetBuilder.InternalFillObject(data, position, false);

            return this;
        }
        
        public TableBuilder Fill(IEnumerable<IEnumerable<object>> data)
        {
            if (!HasHeaders)
                throw new ArgumentException("Table has no header");

            if (HasData)
                throw new ArgumentException("Table already has data");

            HasData = true;
            var position = Position.AddRow(1);

            BodyValue = SheetBuilder.InternalFill(data, position);

            return this;
        }

        public TableBuilder AddHeader(IEnumerable<string> headers)
        {
            if (HasData)
                throw new ArgumentException("Table already has data");

            if (HasHeaders)
                throw new ArgumentException("Table already has header");
            
            HasHeaders = true;
            List<IEnumerable<object>> data = new();
            data.Add(headers);

            HeaderValue = SheetBuilder.InternalFill(data, Position);
            return this;
        }

        #endregion "Builder"

        #region "Generator"

        internal void GenerateTablePart(WorksheetPart worksheetPart)
        {
            if (!HasData)
                return;

            var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();

            Table table = new()
            {
                Id = (uint)SpreadsheetBuilder.Tables.IndexOf(this) + 1,
                Name = TableName,
                DisplayName = TableName,
                TotalsRowShown = false
            };

            var reference = $"{Position}:{BodyValue.FinalPosition}";
            table.Reference = reference;

            AutoFilter autoFilter = new();
            autoFilter.Reference = reference;

            table.Append(autoFilter);

            TableColumns tableColumns = new();
            tableColumns.Count = Position.CountColumns(BodyValue.FinalPosition);
            var headerData = HeaderValue.Value.GetAllData().First().ToList();
                
            for (uint i = 0; i < tableColumns.Count; i++)
            {
                TableColumn tableColumn = new();
                tableColumn.Id = i + 1;
                tableColumn.Name = headerData[(int)i].ToString();

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

            TableParts tableParts = (TableParts)worksheetPart.Worksheet.ChildElements.FirstOrDefault(ce => ce is TableParts);

            if (tableParts == null)
            {
                tableParts = new() { Count = (UInt32)1 };
            }
            else
            {
                if (tableParts.Count.HasValue)
                    tableParts.Count++;
                else
                    tableParts.Count = (UInt32)1;
            }

            TablePart tablePart = new() { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) };

            tableParts.Append(tablePart);

            if (tableParts.Count <= 1)
                worksheetPart.Worksheet.Append(tableParts);
        }

        #endregion "Generator"
    }
}
