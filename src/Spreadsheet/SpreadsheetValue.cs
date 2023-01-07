

using System;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetValue
    {
        internal SpreadsheetValue(SheetPosition position, ReportValue value, SpreadsheetFormatter? formatter = null)
        {
            Position = position;
            FinalPosition = position;
            Value = value;

            if (formatter == null)
            {
                formatter = new()
                {
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

        internal SheetPosition FinalPosition;

        internal bool ProcessRow(uint row)
        {
            if (Processed) return false;

            if (row < Position.Row) return false;

            var hasNextRow = Value.NextRow();

            if (hasNextRow)
                FinalPosition = FinalPosition.AddRow();

            return hasNextRow;
        }

        internal bool ProcessColumn(uint column)
        {
            if (Processed) return false;

            if (column < Position.Column) return false;

            var hasNextColumn = Value.NextColumn();

            if (hasNextColumn)
                FinalPosition = FinalPosition.AddColumn();

            return hasNextColumn;
        }

        internal SpreadsheetFormatter GetFormatter(uint row, uint column)
        {
            return new()
            {
                Format = null,
                FormatProvider = null
            };
        }
    }
}
