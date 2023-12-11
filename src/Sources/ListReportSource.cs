using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace BlackDigital.Report.Sources
{
    public class ListReportSource : ReportSource
    {
        #region "Constructors"

        public ListReportSource() { }

        public ListReportSource(IEnumerable<object> data)
            : this()
        {
            Load(data);
        }

        #endregion "Constructors"

        #region "Properties"

        protected IEnumerable<object>? Data;

        protected bool RowProcessed = false;
        protected int ColumnPosition = -1;

        public override uint RowCount => Data == null ? 0u : 1u;
        public override uint ColumnCount => Data == null ? 0u : (uint)Data.Count();

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? value)
        {
            return value is IEnumerable && value is not string;
        }

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new Exception("Invalid data type");

            Data = ((IEnumerable)data).Cast<object>();
        }

        public override Task<bool> NextRowAsync()
        {
            if (!RowProcessed)
            {
                RowProcessed = true;
                ColumnPosition = -1;
                return Task.FromResult(true);
            }
            else
                return Task.FromResult(false);
        }

        public override Task<bool> NextColumnAsync()
        {
            ColumnPosition++;

            if (ColumnPosition < Data!.Count())
                return Task.FromResult(true);
            else
                return Task.FromResult(false);
        }

        public override Task<object?> GetValueAsync()
        {
            if (!Processed
                && RowProcessed
                && ColumnPosition >= 0
                && ColumnPosition < Data!.Count())
                return Task.FromResult<object?>(Data!.ElementAt(ColumnPosition));
            else
                return Task.FromResult<object?>(null);
        }

        public override Task ResetAsync()
        {
            Processed = false;
            RowProcessed = false;
            ColumnPosition = -1;
            return Task.CompletedTask;
        }

        #endregion "ReportSource"
    }
}
