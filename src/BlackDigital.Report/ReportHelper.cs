using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    internal static class ReportHelper
    {
        internal static List<string> GetObjectHeader<T>()
        {
            var properties = typeof(T).GetProperties();
            return properties.Select(p => p.Name).ToList();
        }
        
        internal static List<List<object>> ObjectToData<T>(IEnumerable<T> data, bool generateHeader = true)
        {
            var list = new List<List<object>>();

            if (generateHeader)
                list.Add(GetObjectHeader<T>().Cast<object>().ToList());
            
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
