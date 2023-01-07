
namespace BlackDigital.Report
{
    internal class SingleReportValue : ReportValue
    {
        internal SingleReportValue(object? value)
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
            if (!Processed)
            {
                return Value;
            }
            else
            {
                return null;
            }
        }
    }
}
