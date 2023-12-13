using System;
using System.Linq;
using System.Resources;
using System.Reflection;
using System.Globalization;
using System.Collections.Generic;
using BlackDigital.Report.Sources;
using System.ComponentModel.DataAnnotations;

namespace BlackDigital.Report
{
    internal static class ReportHelper
    {
        internal static ReportSource GetObjectHeader<T>(ResourceManager? resource, CultureInfo? culture)
        {
            var properties = GetPropertiesAndAttributes<T>();
            var header = properties.Select(p =>
            {
                string columnName = p.Item2?.GetName() ?? p.Item1.Name;

                if (resource != null)
                {
                    var name = resource.GetString(columnName, culture);

                    if (name != null)
                        columnName = name;
                }

                return columnName;
            }).ToList();

            var source = new ListReportSource();
            source.Load(header);

            return source;
        }
        
        internal static ReportSource ObjectToData<T>(IEnumerable<T> data)
        {
            var list = new List<List<object>>();
            
            var properties = GetPropertiesAndAttributes<T>();

            foreach (var row in data)
            {
                var dataRow = new List<object>();

                foreach (var property in properties)
                {
                    dataRow.Add(property.Item1?.GetValue(row) ?? string.Empty);
                }

                list.Add(dataRow);
            }

            var source = new EnumerableReportSource();
            source.Load(list);

            return source;
        }

        internal static List<Tuple<PropertyInfo, DisplayAttribute?>> GetPropertiesAndAttributes<T>()
        {
            var properties = typeof(T).GetProperties();

            if (properties == null)
                return new();

            var all = properties.Select<PropertyInfo, Tuple<PropertyInfo, DisplayAttribute?>>(property => new(
                            property,
                            property?.GetCustomAttributes(typeof(DisplayAttribute), false)
                                     .Cast<DisplayAttribute>()
                                     .FirstOrDefault()
            )).ToList();

            all.RemoveAll(a => a.Item2 != null && a.Item2.GetAutoGenerateField() == false);

            return all.OrderBy(a => a.Item2?.Order ?? Int32.MaxValue).ToList();
        }
    }
}
