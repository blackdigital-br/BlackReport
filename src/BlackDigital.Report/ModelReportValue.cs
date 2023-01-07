using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace BlackDigital.Report
{
    internal class ModelReportValue<T> : ReportValue
    {
        internal ModelReportValue(IEnumerable<T> modelList, bool includePropertyNames = true)
        {
            IncludePropertyNames = includePropertyNames;
            ModelList = modelList;
            Properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                  .ToList();

            PropertyNamesProcessed = IncludePropertyNames == false;
        }

        private readonly bool IncludePropertyNames;
        private readonly IEnumerable<T> ModelList;
        protected readonly IEnumerable<PropertyInfo> Properties;
        protected T CurrentModel;
        private bool PropertyNamesProcessed = false;

        internal override bool NextRow()
        {
            if (PropertyNamesProcessed)
            {
                if (ModelList.GetEnumerator().MoveNext())
                {
                    CurrentModel = ModelList.GetEnumerator().Current;
                    Properties.GetEnumerator().Reset();
                    return true;
                }
                else
                {
                    Processed = true;
                    return false;
                }
            }
            else
            {
                PropertyNamesProcessed = true;
                return true;
            }
        }

        internal override bool NextColumn()
        {
            return Properties.GetEnumerator().MoveNext();
        }

        internal override object? GetValue()
        {
            if (!Processed)
            {
                if (PropertyNamesProcessed)
                {
                    return Properties.GetEnumerator().Current.GetValue(CurrentModel);
                }
                else
                {
                    return Properties.GetEnumerator().Current.Name;
                }
            }
            else
            {
                return null;
            }
        }
    }
}
