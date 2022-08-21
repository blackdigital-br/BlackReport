using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Resources;
using System.Globalization;
using DocumentFormat.OpenXml.CustomProperties;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace BlackDigital.Report
{
    internal static class ReportHelper
    {
        internal static List<string> GetObjectHeader<T>(ResourceManager? resource, CultureInfo? culture)
        {
            var properties = GetPropertiesAndAttributes<T>();
            return properties.Select(p => {
                string columnName = p.Item2?.GetName() ?? p.Item1.Name;

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


            return list;
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
