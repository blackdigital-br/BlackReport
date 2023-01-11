using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    internal class TemplateReportSource : ReportSource
    {
        public TemplateReportSource(string name)
        {
            Name = name;
        }

        private readonly string Name;
        public ReportSource? Source { get; set; }

        internal void SetSource(Dictionary<string, ReportSource> sources)
        {
            if (sources.ContainsKey(Name))
                Source = sources[Name];
        }

        internal override bool NextColumn()
        {
            return Source?.NextColumn() ?? false;
        }

        internal override bool NextRow()
        {
            bool result = Source?.NextRow() ?? false;
            Source = null;

            return result;
        }

        internal override object? GetValue()
        {
            return Source?.GetValue();
        }

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            return Source?.GetAllData() ?? Enumerable.Empty<IEnumerable<object>>();
        }
    }
}
