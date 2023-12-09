using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace BlackDigital.Report.Sources
{
    public class DataTableReportSource : ReportSource
    {
        #region "Constructors"

        public DataTableReportSource() { }

        public DataTableReportSource(DataTable table)
            : this()
        {
            Load(table);
        }

        #endregion "Constructors"

        #region "Properties"

        protected DataTable? Table;
        private IEnumerator<DataRow>? RowEnumarator;
        private IEnumerator? ColumnEnumarator;
        protected DataRow? CurrentRow;

        private bool Row = false;
        private bool Column = false;

        private uint _rowCount = 0;
        private uint _columnCount = 0;
        private bool _columnCounted = false;

        public override uint RowCount => _rowCount;
        public override uint ColumnCount => _columnCount;

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? data)
            => data is DataTable;

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new Exception("Invalid data type");

            Table = (DataTable)data;
            RowEnumarator = Table.Rows.Cast<DataRow>().GetEnumerator();
            RowEnumarator.Reset();
        }

        public override bool NextRow()
        {
            if (RowEnumarator == null)
                throw new System.Exception("ReportSource is not loaded");

            if (RowEnumarator.MoveNext())
            {
                _rowCount++;
                CurrentRow = RowEnumarator.Current;
                ColumnEnumarator = CurrentRow.ItemArray.GetEnumerator();
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

            Column = ColumnEnumarator.MoveNext();

            if (Column && !_columnCounted)
                _columnCount++;
            else
                _columnCounted = true;

            return Column;
        }

        public override object? GetValue()
        {
            if (ColumnEnumarator == null)
                throw new System.Exception("ReportSource is not loaded");

            if (!Processed && Column && Row)
            {
                return ColumnEnumarator.Current;
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
        }

        #endregion "ReportSource"
    }
}
