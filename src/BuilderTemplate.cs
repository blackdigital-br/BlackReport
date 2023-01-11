using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public static class BuilderTemplate
    {
        private static Dictionary<string, ReportBuilder> Templates = new();

        internal static void AddTemplate(string name, ReportBuilder builder)
        {
            Templates.Add(name, builder);
        }
    }
}
