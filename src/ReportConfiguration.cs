using System;
using System.Collections.Generic;
using BlackDigital.Report.Sources;
using BlackDigital.Report.Spreadsheet;


namespace BlackDigital.Report
{
    public class ReportConfiguration
    {
        public ReportConfiguration()
        {
            Configurations = new()
            {
                { "Sources", new SourceConfiguration() },
                { "Spreadsheet", new SpreadsheetConfiguration() }
            };
        }

        private Dictionary<string, object> Configurations { get; set; }

        public SourceConfiguration Sources
            => Get<SourceConfiguration>("Sources");

        public SpreadsheetConfiguration Spreadsheet
            => Get<SpreadsheetConfiguration>("Spreadsheet");


        public void Add(string name, object configuration)
        {
            if (Configurations.ContainsKey(name))
                throw new Exception($"Configuration {name} already exists");
            else
                Configurations.Add(name, configuration);
        }

        public bool Contains(string name)
            => Configurations.ContainsKey(name);

        public T Get<T>(string name)
        {
            if (Configurations.ContainsKey(name))
                return (T)Configurations[name];
            else
                throw new Exception($"Configuration {name} not found");
        }
    }
}
