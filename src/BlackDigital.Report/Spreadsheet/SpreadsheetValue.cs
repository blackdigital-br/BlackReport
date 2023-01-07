
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace BlackDigital.Report.Spreadsheet
{
    internal class SpreadsheetValue
    {
        public SpreadsheetValue(SheetPosition position, ReportValue value, SpreadsheetFormatter? formatter = null)
        {
            Position = position;
            Value = value;

            if (formatter == null)
            {
                formatter = new()
                {
                    CellReference = SpreadsheetHelper.NumberToExcelColumn(position.Row, position.Column),
                    Format = null,
                    FormatProvider = null,
                    Value = value,
                    ValueType = value?.GetType() ?? typeof(object)
                };
            }
            
            Formatter = formatter;
        }

        public readonly SheetPosition Position;

        public readonly ReportValue Value;

        public readonly SpreadsheetFormatter? Formatter;

        internal bool Processed => Value.Processed;

        internal bool ProcessRow(uint row)
        {
            if (Processed) return false;

            if (row < Position.Row) return false;
            
            return Value.NextRow();
        }

        internal bool ProcessColumn(uint column)
        {
            if (Processed) return false;

            if (column < Position.Column) return false;

            return Value.NextColumn();
        }
    }
}
