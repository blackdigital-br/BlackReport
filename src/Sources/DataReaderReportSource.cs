using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Runtime.Versioning;

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

        public override bool NextRow()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            Row = Reader.Read();
            Columns = Reader.FieldCount;
            Column = 0;

            if (Row)
            {
                _rowCount++;
            }
            else
            {
                Processed = true;
                Reader.Close();
            }

            return Row;
        }

        public override bool NextColumn()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            if (Column < Columns)
            {
                Column++;
                return true;
            }
            else
            {
                return false;
            }
        }

        public override object? GetValue()
        {
            if (Reader == null)
                throw new Exception("ReportSource is not loaded");

            if (!Processed
                && Row
                && Column <= Columns)
            {
                return Reader.GetValue(Column - 1);
            }
            else
            {
                return null;
            }
        }

        public override void Reset()
        {
            throw new Exception("DataReaderReportSource cannot be reset");
        }

        #endregion "ReportSource"   
    }
}
