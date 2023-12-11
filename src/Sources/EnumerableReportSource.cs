using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace BlackDigital.Report.Sources
{
    public class EnumerableReportSource : ReportSource
    {
        #region "Constructors"

        public EnumerableReportSource() { }

        public EnumerableReportSource(IEnumerable<IEnumerable<object>> data)
            : this()
        {
            Load(data);
        }

        #endregion "Constructors"

        #region "Properties"

        protected IEnumerable<IEnumerable<object>>? Data;
        protected IEnumerable<object>? CurrentRow;

        protected IEnumerator<IEnumerable<object>>? RowEnumarator;
        protected IEnumerator<object>? ColumnEnumarator;

        private uint _rowCount = 0;
        private uint _columnCount = 0;
        private bool _columnCounted = false;

        public override uint RowCount => _rowCount;
        public override uint ColumnCount => _columnCount;

        private bool Row = false;
        private bool Column = false;

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? data)
            => data is IEnumerable<IEnumerable<object>>;

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new System.Exception("Invalid data type");

            Data = (IEnumerable<IEnumerable<object>>)data;
            RowEnumarator = Data.GetEnumerator();
            RowEnumarator.Reset();
        }

        public override Task<bool> NextRowAsync()
        {
            if (RowEnumarator == null)
                throw new Exception("ReportSource is not loaded");

            if (RowEnumarator.MoveNext())
            {
                _rowCount++;
                CurrentRow = RowEnumarator.Current;
                ColumnEnumarator = CurrentRow.GetEnumerator();
                ColumnEnumarator.Reset();
                Row = true;
                return Task.FromResult(true);
            }
            else
            {
                Row = false;
                Processed = true;
                return Task.FromResult(false);
            }
        }

        public override Task<bool> NextColumnAsync()
        {
            if (ColumnEnumarator == null)
                throw new Exception("ReportSource is not loaded");

            Column = ColumnEnumarator.MoveNext();

            if (Column && !_columnCounted)
                _columnCount++;
            else
                _columnCounted = true;

            return Task.FromResult(Column);
        }

        public override Task<object?> GetValueAsync()
        {
            if (Data == null)
                throw new Exception("ReportSource is not loaded");

            if (!Processed && Column && Row)
            {
                return Task.FromResult(ColumnEnumarator?.Current);
            }
            else
            {
                return Task.FromResult<object?>(null);
            }
        }

        public override Task ResetAsync()
        {
            RowEnumarator?.Reset();
            ColumnEnumarator?.Reset();
            Row = false;
            Column = false;
            return Task.CompletedTask;
        }

        #endregion "ReportSource"
    }
}
