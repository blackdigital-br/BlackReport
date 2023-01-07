
using System.Collections.Generic;

namespace BlackDigital.Report
{
    internal class SingleReportSource : ReportSource
    {
        internal SingleReportSource(object? value)
        {
            Value = value;
        }

        protected bool RowProcessed = false;
        protected bool ColumnProcessed = false;
        internal readonly object? Value;

        internal override bool NextRow()
        {
            if (!RowProcessed)
            {
                RowProcessed = true;
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
            if (!ColumnProcessed)
            {
                ColumnProcessed = true;
                return true;
            }
            else
            {
                Processed = true;
                return false;
            }
        }

        internal override object? GetValue()
        {

            if (!Processed
                && RowProcessed
                && ColumnProcessed)
            {
                return Value;
            }
            else
            {
                return null;
            }
        }

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            return new List<IEnumerable<object>>()
            {
                new object[] { Value }
            };
        }
    }
}
