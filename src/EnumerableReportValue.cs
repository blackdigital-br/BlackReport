using System.Collections.Generic;

namespace BlackDigital.Report
{
    internal class EnumerableReportValue : ReportValue
    {
        internal EnumerableReportValue(IEnumerable<IEnumerable<object>> data)
        {
            Data = data;
            RowEnumarator = Data.GetEnumerator();
            RowEnumarator.Reset();
        }

        private readonly IEnumerable<IEnumerable<object>> Data;
        protected IEnumerable<object> CurrentRow;

        private IEnumerator<IEnumerable<object>> RowEnumarator;
        private IEnumerator<object> ColumnEnumarator;

        internal override bool NextRow()
        {
            if (RowEnumarator.MoveNext())
            {
                CurrentRow = RowEnumarator.Current;
                ColumnEnumarator = CurrentRow.GetEnumerator();
                ColumnEnumarator.Reset();
                return true;
            }
            else
            {
                Processed = true;
                return false;
            }
        }

        internal override bool NextColumn()
        {
            return ColumnEnumarator.MoveNext();
        }

        internal override object? GetValue()
        {
            if (!Processed)
            {
                return ColumnEnumarator.Current;
            }
            else
            {
                return null;
            }
        }
    }
}
