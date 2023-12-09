using System;

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

        public abstract bool NextRow();

        public abstract bool NextColumn();

        public abstract object? GetValue();

        public abstract void Reset();
    }
}
