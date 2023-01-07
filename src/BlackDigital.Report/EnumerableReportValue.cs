using System.Collections.Generic;

namespace BlackDigital.Report
{
    internal class EnumerableReportValue : ReportValue
    {
        internal EnumerableReportValue(IEnumerable<IEnumerable<object>> data)
        {
            Data = data;
        }

        private readonly IEnumerable<IEnumerable<object>> Data;
        protected IEnumerable<object> CurrentRow;

        internal override bool NextRow()
        {
            if (Data.GetEnumerator().MoveNext())
            {
                CurrentRow = Data.GetEnumerator().Current;
                CurrentRow.GetEnumerator().Reset();
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
            return CurrentRow.GetEnumerator().MoveNext();
        }

        internal override object? GetValue()
        {
            if (!Processed)
            {
                return CurrentRow.GetEnumerator().Current;
            }
            else
            {
                return null;
            }
        }
    }
}
