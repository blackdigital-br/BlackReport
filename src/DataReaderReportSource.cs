using System;
using System.Collections.Generic;
using System.Data.Common;

namespace BlackDigital.Report
{
    internal class DataReaderReportSource : ReportSource
    {
        internal DataReaderReportSource(DbDataReader reader)
        {
            Reader = reader;
        }

        private readonly DbDataReader Reader;
        private bool Row;
        private int Column;
        private int Columns;

        internal override bool NextRow()
        {
            Row = Reader.Read();
            Columns = Reader.FieldCount;
            Column = 0;

            if (!Row)
            {
                Processed = true;
                Reader.Close();
            }

            return Row;
        }

        internal override bool NextColumn()
        {
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

        internal override object? GetValue()
        {
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

        internal override IEnumerable<IEnumerable<object>> GetAllData()
        {
            throw new NotImplementedException();
        }
    }
}
