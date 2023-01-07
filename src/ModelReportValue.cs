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
                                  .AsEnumerable();

            PropertyNamesProcessed = IncludePropertyNames == false;
            RowEnumarator = ModelList.GetEnumerator();
            RowEnumarator.Reset();
        }

        private readonly bool IncludePropertyNames;
        private readonly IEnumerable<T> ModelList;
        protected readonly IEnumerable<PropertyInfo> Properties;
        protected T? CurrentModel;
        private bool PropertyNamesProcessed = false;

        private IEnumerator<T> RowEnumarator;
        private IEnumerator<PropertyInfo>? ColumnEnumarator;

        internal override bool NextRow()
        {
            if (PropertyNamesProcessed)
            {
                if (RowEnumarator.MoveNext())
                {
                    CurrentModel = RowEnumarator.Current;
                    ColumnEnumarator = Properties.GetEnumerator();
                    ColumnEnumarator.Reset();
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
                ColumnEnumarator = Properties.GetEnumerator();
                ColumnEnumarator.Reset();
                return true;
            }
        }

        internal override bool NextColumn()
        {
            var hasColumn = ColumnEnumarator?.MoveNext() ?? false;

            if (!hasColumn && !PropertyNamesProcessed)
                PropertyNamesProcessed = true;

            return hasColumn;
        }

        internal override object? GetValue()
        {
            if (!Processed)
            {
                if (PropertyNamesProcessed)
                {
                    if (CurrentModel == null)
                        return null;

                    return ColumnEnumarator?.Current.GetValue(CurrentModel);
                }
                else
                {
                    return ColumnEnumarator?.Current.Name;
                }
            }
            else
            {
                return null;
            }
        }
    }
}
