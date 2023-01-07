

using System;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetValue
    {
        internal SpreadsheetValue(SheetPosition position, ReportSource value, ValueFormatter? formatter = null)
        {
            Position = position;
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

        internal readonly ReportSource Value;

        public readonly ValueFormatter? Formatter;

        internal bool Processed => Value.Processed;

        private uint RowProcessed = 0;
        private uint ColumnProcessed = 0;

        internal SheetPosition FinalPosition
        {
            get
            {
                return Position.Add(ColumnProcessed - 1, RowProcessed - 1);
            }
        }

        internal bool ProcessRow(uint row)
        {
            if (Processed) return false;

            if (row < Position.Row) return false;

            var hasNextRow = Value.NextRow();

            if (hasNextRow)
            {
                RowProcessed++;
            }

            return hasNextRow;
        }

        internal bool ProcessColumn(uint column)
        {
            if (Processed) return false;

            if (column < Position.Column) return false;

            var hasNextColumn = Value.NextColumn();

            if (hasNextColumn
                && RowProcessed == 1)
                ColumnProcessed++;

            return hasNextColumn;
        }

        internal ValueFormatter GetFormatter(uint row, uint column)
        {
            return new()
            {
                Format = null,
                FormatProvider = null
            };
        }
    }
}
