using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    internal sealed class DefaultCellCreate : ICellCreate
    {
        private static Dictionary<Type, Func<SpreadsheetFormatter, Cell>> _cellCreators = new()
        {
            { typeof(bool), CreateCellBoolean },
            { typeof(char), CreateCellString },
            { typeof(string), CreateCellString },
            { typeof(byte), CreateCellNumber },
            { typeof(sbyte), CreateCellNumber },
            { typeof(ushort), CreateCellNumber },
            { typeof(short), CreateCellNumber },
            { typeof(int), CreateCellNumber },
            { typeof(uint), CreateCellNumber },
            { typeof(ulong), CreateCellNumber },
            { typeof(long), CreateCellNumber },
            { typeof(float), CreateCellNumber },
            { typeof(double), CreateCellNumber },
            { typeof(decimal), CreateCellNumber },
            { typeof(DateTime), CreateCellDateTime },
            { typeof(DateTimeOffset), CreateCellDateTime },
            { typeof(TimeSpan), CreateCellTimespan }
        };
            
        public Cell CreateCell(SpreadsheetFormatter formatter)
        {
            if (formatter.Value == null)
                return CreateCellDefault(formatter);

            if (_cellCreators.ContainsKey(formatter.ValueType))
                return _cellCreators[formatter.ValueType](formatter);

            Type? nullType = Nullable.GetUnderlyingType(formatter.ValueType);
            if (nullType != null && _cellCreators.ContainsKey(nullType))
                return _cellCreators[nullType](formatter);

            return CreateCellDefault(formatter);
        }

        private static Cell CreateCellDefault(SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(string.Empty),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellString(SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(formatter.Value?.ToString() ?? string.Empty),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellBoolean(SpreadsheetFormatter formatter)
        {   
            return new Cell()
            {
                DataType = CellValues.Boolean,
                CellValue = new CellValue(Convert.ToBoolean(formatter.Value)),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellNumber(SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(Convert.ToDouble(formatter.Value)),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellDateTime(SpreadsheetFormatter formatter)
        {
            DateTime value;

            if (formatter.Value is DateTime dateTime)
                value = dateTime;
            else if (formatter.Value is DateTimeOffset dateTimeOffset)
                value = dateTimeOffset.DateTime;
            else
                throw new InvalidOperationException("Invalid value type");

            return new Cell()
            {
                DataType = CellValues.Date,
                CellValue = new CellValue(value),
                CellReference = formatter.CellReference,
                StyleIndex = 2u
            };
        }

        private static Cell CreateCellTimespan(SpreadsheetFormatter formatter)
        {
            TimeSpan value;

            if (formatter.Value is TimeSpan timeSpan)
                value = timeSpan;
            else
                throw new InvalidOperationException("Invalid value type");

            return new Cell()
            {
                DataType = CellValues.Number,
                //CellValue = new CellValue((new DateTime(1900, 1, 1)).Add(value)),
                CellValue = new CellValue(value.TotalSeconds / 86400),
                CellReference = formatter.CellReference,
                StyleIndex = 1u
            };
        }
    }
}
