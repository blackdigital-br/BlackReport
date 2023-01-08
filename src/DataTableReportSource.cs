using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace BlackDigital.Report
{
    internal class DataTableReportSource : ReportSource
    {
        internal DataTableReportSource(DataTable table)
        {
            Table = table;
            RowEnumarator = Table.Rows.Cast<DataRow>().GetEnumerator();
            RowEnumarator.Reset();
        }

        private readonly DataTable Table;
        private IEnumerator<DataRow> RowEnumarator;
        private IEnumerator ColumnEnumarator;

        protected DataRow CurrentRow;

        private bool Row = false;
        private bool Column = false;

        internal override bool NextRow()
        {
            if (RowEnumarator.MoveNext())
            {
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

        internal override bool NextColumn()
        {
            Column = ColumnEnumarator.MoveNext();
            return Column;
        }

        internal override object? GetValue()
        {
            if (!Processed && Column && Row)
            {
                return ColumnEnumarator.Current;
            }
            else
            {
                return null;
            }
        }

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            return Table.Rows.Cast<DataRow>().Select(r => r.ItemArray.Cast<object>());
        }
    }
}
