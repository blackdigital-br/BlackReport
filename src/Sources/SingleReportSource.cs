using System;
using System.Threading.Tasks;


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
            return (data is string || type.IsValueType)
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

        public override Task<bool> NextRowAsync()
        {
            if (!RowProcessed)
            {
                RowProcessed = true;
                _rowCount++;
                return Task.FromResult(true);
            }
            else
            {
                Processed = true;
                return Task.FromResult(false);
            }
        }

        public override Task<bool> NextColumnAsync()
        {
            if (!ColumnProcessed)
            {
                ColumnProcessed = true;
                _columnCount++;
                return Task.FromResult(true);
            }
            else
            {
                Processed = true;
                return Task.FromResult(false);
            }
        }

        public override Task<object?> GetValueAsync()
        {

            if (!Processed
                && RowProcessed
                && ColumnProcessed)
            {
                return Task.FromResult(Value);
            }
            else
            {
                return Task.FromResult<object?>(null);
            }
        }

        public override Task ResetAsync()
        {
            Processed = false;
            RowProcessed = false;
            ColumnProcessed = false;
            return Task.CompletedTask;
        }

        #endregion "ReportSource"
    }
}
