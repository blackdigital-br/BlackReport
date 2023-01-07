

using System;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetValue
    {
        internal SpreadsheetValue(SheetPosition position, ReportValue value, SpreadsheetFormatter? formatter = null)
        {
            Position = position;
            Value = value;

            if (formatter == null)
            {
                formatter = new()
                {
                    CellReference = SpreadsheetHelper.NumberToExcelColumn(position.Row, position.Column),
                    Format = null,
                    FormatProvider = null
                };
            }
            
            Formatter = formatter;
        }

        public readonly SheetPosition Position;

        internal readonly ReportValue Value;

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

        internal SpreadsheetFormatter GetFormatter(uint row, uint column)
        {
            return new()
            {
                CellReference = SpreadsheetHelper.NumberToExcelColumn(row, column),
                Format = null,
                FormatProvider = null
            };
        }
    }
}
