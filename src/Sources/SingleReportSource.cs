using System;


namespace BlackDigital.Report.Sources
{
    public class SingleReportSource : ReportSource
    {
        #region "Constructors"

        public SingleReportSource() { }

        public SingleReportSource(object data)
            : this()
        {
            Load(data);
        }

        #endregion "Constructors"

        #region "Properties"

        protected bool RowProcessed = false;
        protected bool ColumnProcessed = false;
        protected object? Value;

        private uint _rowCount = 0;
        private uint _columnCount = 0;

        public override uint RowCount => _rowCount;
        public override uint ColumnCount => _columnCount;

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? data)
        {
            return (
                data is string
                || type.IsValueType
                )
                && data != null;
        }

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new Exception("Invalid data type");

            Value = data;
        }

        public override bool NextRow()
        {
            if (!RowProcessed)
            {
                RowProcessed = true;
                _rowCount++;
                return true;
            }
            else
            {
                Processed = true;
                return false;
            }
        }

        public override bool NextColumn()
        {
            if (!ColumnProcessed)
            {
                ColumnProcessed = true;
                _columnCount++;
                return true;
            }
            else
            {
                Processed = true;
                return false;
            }
        }

        public override object? GetValue()
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

        public override void Reset()
        {
            Processed = false;
            RowProcessed = false;
            ColumnProcessed = false;
        }

        #endregion "ReportSource"
    }
}
