
using System.Collections.Generic;

namespace BlackDigital.Report
{
    internal abstract class ReportSource
    {
        internal ReportSource()
        {
        }
        
        public bool Processed { get; protected set; } = false;

        internal abstract bool NextRow();

        internal abstract bool NextColumn();

        internal abstract object? GetValue();

        internal abstract IEnumerable<IEnumerable<object>> GetAllData();
    }
}
