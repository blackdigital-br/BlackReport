using System;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public abstract class ReportSource
    {
        public ReportSource()
        {
        }
        
        public bool Processed { get; protected set; } = false;

        public abstract uint RowCount { get; }

        public abstract uint ColumnCount { get; }

        public abstract bool IsSourceType(Type type, object? value);
        
        public abstract void Load(object data);

        public abstract Task<bool> NextRowAsync();

        public abstract Task<bool> NextColumnAsync();

        public abstract Task<object?> GetValueAsync();

        public abstract Task ResetAsync();
    }
}
