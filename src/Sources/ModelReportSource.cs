using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;

namespace BlackDigital.Report.Sources
{
    public class ModelReportSource<T> : ReportSource
    {
        #region "Constructors"

        public ModelReportSource() { }

        public ModelReportSource(IEnumerable<T> data)
            : this()
        {
            Load(data);
        }

        #endregion "Constructors"

        #region "Properties"

        protected IEnumerable<T>? ModelList;
        protected IEnumerable<PropertyInfo>? Properties;
        protected T? CurrentModel;

        private IEnumerator<T>? RowEnumarator;
        private IEnumerator<PropertyInfo>? ColumnEnumarator;

        private bool Row = false;
        private bool Column = false;

        private uint _rowCount = 0;
        private uint _columnCount = 0;
        private bool columnCounted = false;

        public override uint RowCount => _rowCount;
        public override uint ColumnCount => _columnCount;

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? data)
            => data is IEnumerable<T>;

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new System.Exception("Invalid data type");

            ModelList = (IEnumerable<T>)data;
            Properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                  .AsEnumerable();

            RowEnumarator = ModelList.GetEnumerator();
            RowEnumarator.Reset();
        }

        public override bool NextRow()
        {
            if (RowEnumarator == null
                || Properties == null)
                throw new System.Exception("ReportSource is not loaded");

            if (RowEnumarator.MoveNext())
            {
                _rowCount++;
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

        public override bool NextColumn()
        {
            if (ColumnEnumarator == null)
                throw new System.Exception("ReportSource is not loaded");

            Column = ColumnEnumarator?.MoveNext() ?? false;

            if (Column && !columnCounted)
                _columnCount++;
            else
                columnCounted = true;

            return Column;
        }

        public override object? GetValue()
        {
            if (ModelList == null)
                throw new System.Exception("ReportSource is not loaded");

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

        public override void Reset()
        {
            RowEnumarator?.Reset();
            ColumnEnumarator?.Reset();
            Row = false;
            Column = false;
            Processed = false;
        }

        #endregion "ReportSource"
    }
}
