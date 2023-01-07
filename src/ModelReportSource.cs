using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace BlackDigital.Report
{
    internal class ModelReportSource<T> : ReportSource
    {
        internal ModelReportSource(IEnumerable<T> modelList)
        {
            ModelList = modelList;
            Properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                  .AsEnumerable();

            RowEnumarator = ModelList.GetEnumerator();
            RowEnumarator.Reset();
        }

        private readonly IEnumerable<T> ModelList;
        protected readonly IEnumerable<PropertyInfo> Properties;
        protected T? CurrentModel;

        private IEnumerator<T> RowEnumarator;
        private IEnumerator<PropertyInfo>? ColumnEnumarator;

        private bool Row = false;
        private bool Column = false;

        internal override bool NextRow()
        {
            if (RowEnumarator.MoveNext())
            {
                CurrentModel = RowEnumarator.Current;
                ColumnEnumarator = Properties.GetEnumerator();
                ColumnEnumarator.Reset();
                Row = true;
                return true;
            }
            else
            {
                Row = false;
                Processed = true;
                return false;
            }
        }

        internal override bool NextColumn()
        {
            Column = ColumnEnumarator?.MoveNext() ?? false;
            return Column;
        }

        internal override object? GetValue()
        {
            if (!Processed && Column && Row)
            {
                if (CurrentModel == null)
                    return null;

                return ColumnEnumarator?.Current.GetValue(CurrentModel);
            }
            else
            {
                return null;
            }
        }

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            return ReportHelper.ObjectToData(ModelList);
        }
    }
}
