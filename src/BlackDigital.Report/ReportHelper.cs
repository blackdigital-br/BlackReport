using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Resources;
using System.Globalization;

namespace BlackDigital.Report
{
    internal static class ReportHelper
    {
        internal static List<string> GetObjectHeader<T>(ResourceManager? resource, CultureInfo? culture)
        {
            var properties = typeof(T).GetProperties();
            return properties.Select(p => {
                string columnName = p.Name;
                
                if (resource != null)
                {
                    var name = resource.GetString(columnName, culture);

                    if (name != null)
                        columnName = name;
                }
                
                return columnName;
            }).ToList();
        }
        
        internal static List<List<object>> ObjectToData<T>(IEnumerable<T> data, bool generateHeader = true, ResourceManager? resource = null, CultureInfo? culture = null)
        {
            var list = new List<List<object>>();

            if (generateHeader)
                list.Add(GetObjectHeader<T>(resource, culture)
                                        .Cast<object>().ToList());
            
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
