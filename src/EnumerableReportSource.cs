using System.Collections.Generic;

namespace BlackDigital.Report
{
    internal class EnumerableReportSource : ReportSource
    {
        internal EnumerableReportSource(IEnumerable<IEnumerable<object>> data)
        {
            Data = data;
            RowEnumarator = Data.GetEnumerator();
            RowEnumarator.Reset();
        }

        private readonly IEnumerable<IEnumerable<object>> Data;
        protected IEnumerable<object> CurrentRow;

        private IEnumerator<IEnumerable<object>> RowEnumarator;
        private IEnumerator<object> ColumnEnumarator;

        private bool Row = false;
        private bool Column = false;

        internal override bool NextRow()
        {
            if (RowEnumarator.MoveNext())
            {
                CurrentRow = RowEnumarator.Current;
                ColumnEnumarator = CurrentRow.GetEnumerator();
                ColumnEnumarator.Reset();
                Row = true;
                return true;
            }
            else
            {
                Row = false;
                Processed = true;
                return false;
            }
        }

        internal override bool NextColumn()
        {
            Column = ColumnEnumarator.MoveNext();
            return Column;
        }

        internal override object? GetValue()
        {
            if (!Processed && Column && Row)
            {
                return ColumnEnumarator.Current;
            }
            else
            {
                return null;
            }
        }

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            return Data;
        }
    }
}
