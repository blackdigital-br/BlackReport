using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public class ReportArguments
    {
        protected ReportArguments(IReport report)
        {
            Report = report;
        }

        protected readonly IReport Report;

        
        public Task<byte[]> GenerateAsync(IEnumerable<IEnumerable<object>> data, IEnumerable<string> columns)
        {
            return Report.GenerateReportAsync(data, columns, this);
        }

        public Task<byte[]> GenerateAsync<T>(IEnumerable<T> data)
        {
            return GenerateAsync(ObjectToData(data), typeof(T).GetProperties().Select(p => p.Name));
        }

        protected IEnumerable<IEnumerable<object>> ObjectToData<T>(IEnumerable<T> data)
        {
            var list = new List<IEnumerable<object>>();
            var properties = typeof(T).GetProperties();

            foreach (var row in data)
            {
                var dataRow = new List<object>();

                foreach (var property in properties)
                {
                    dataRow.Add(property.GetValue(row));
                }

                list.Add(dataRow);
            }
            

            return list;
        }
    }
}
