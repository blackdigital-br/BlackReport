using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Runtime.Versioning;
using System.Threading.Tasks;

namespace BlackDigital.Report.Sources
{
    public class DataReaderReportSource : ReportSource
    {
        #region "Constructors"

        public DataReaderReportSource() { }

        public DataReaderReportSource(DbDataReader data)
            : this()
        {
            Load(data);
        }

        #endregion "Constructors"

        #region "Properties"

        private DbDataReader? Reader;
        private bool Row;
        private int Column;
        private int Columns;

        private uint _rowCount = 0;

        public override uint RowCount => _rowCount;
        public override uint ColumnCount => (uint)Columns;

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? data)
            => data is DbDataReader;

        public override void Load(object data)
        {
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new Exception("Invalid data type");

            Reader = (DbDataReader)data;
        }

        public override async Task<bool> NextRowAsync()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            Row = await Reader.ReadAsync();
            Columns = Reader.FieldCount;
            Column = 0;

            if (Row)
            {
                _rowCount++;
            }
            else
            {
                Processed = true;
                await Reader.CloseAsync();
            }

            return Row;
        }

        public override Task<bool> NextColumnAsync()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            if (Column < Columns)
            {
                Column++;
                return Task.FromResult(true);
            }
            else
            {
                return Task.FromResult(false);
            }
        }

        public override Task<object?> GetValueAsync()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            if (!Processed
                && Row
                && Column <= Columns)
            {
                return Task.FromResult<object?>(Reader.GetValue(Column - 1));
            }
            else
            {
                return Task.FromResult<object?>(null);
            }
        }

        public override Task ResetAsync()
        {
            throw new Exception("DataReaderReportSource cannot be reset");
        }

        #endregion "ReportSource"   
    }
}
