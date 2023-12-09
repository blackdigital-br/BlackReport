
using System;

namespace BlackDigital.Report
{
    public class Report
    {
        public Report()
        {
            Configuration = new();
        }

        public ReportConfiguration Configuration { get; private set; }

        public void Configure(Func<ReportConfiguration, ReportConfiguration> action)
        {
            Configuration = action(new());
        }
    }
}
