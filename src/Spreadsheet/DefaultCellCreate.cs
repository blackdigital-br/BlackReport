using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace BlackDigital.Report.Spreadsheet
{
    internal sealed class DefaultCellCreate : ICellCreate
    {
        private static Dictionary<Type, Func<object?, SpreadsheetFormatter, Cell>> _cellCreators = new()
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

#if NET6_0_OR_GREATER
            ,
            { typeof(TimeOnly), CreateCellTimeOnly },
            { typeof(DateOnly), CreateCellDateOnly }
#endif

        };
            
        public Cell CreateCell(object? value, SpreadsheetFormatter formatter)
        {
            if (value == null)
                return CreateCellDefault(value, formatter);

            var valueType = value?.GetType() ?? typeof(object);
            
            if (_cellCreators.ContainsKey(valueType))
                return _cellCreators[valueType](value, formatter);

            Type? nullType = Nullable.GetUnderlyingType(valueType);
            if (nullType != null && _cellCreators.ContainsKey(nullType))
                return _cellCreators[nullType](value, formatter);

            return CreateCellDefault(value, formatter);
        }

        private static Cell CreateCellDefault(object? value, SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(string.Empty),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellString(object? value, SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(value?.ToString() ?? string.Empty),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellBoolean(object? value, SpreadsheetFormatter formatter)
        {   
            return new Cell()
            {
                DataType = CellValues.Boolean,
                CellValue = new CellValue(Convert.ToBoolean(value)),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellNumber(object? value, SpreadsheetFormatter formatter)
        {
            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(Convert.ToDouble(value)),
                CellReference = formatter.CellReference
            };
        }

        private static Cell CreateCellDateTime(object? value, SpreadsheetFormatter formatter)
        {
            DateTime realValue;

            if (value is DateTime dateTime)
                realValue = dateTime;
            else if (value is DateTimeOffset dateTimeOffset)
                realValue = dateTimeOffset.DateTime;
            else
                throw new InvalidOperationException("Invalid value type");

            return new Cell()
            {
                DataType = CellValues.Date,
                CellValue = new CellValue(realValue),
                CellReference = formatter.CellReference,
                StyleIndex = (uint)SpreadsheetFormat.DateTime
            };
        }

        private static Cell CreateCellTimespan(object? value, SpreadsheetFormatter formatter)
        {
            TimeSpan realValue;

            if (value is TimeSpan timeSpan)
                realValue = timeSpan;
            else
                throw new InvalidOperationException("Invalid value type");

            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(realValue.TotalSeconds / 86400),
                CellReference = formatter.CellReference,
                StyleIndex = (uint)SpreadsheetFormat.TimeSpan
            };
        }

#if NET6_0_OR_GREATER

        private static Cell CreateCellTimeOnly(object? value, SpreadsheetFormatter formatter)
        {
            TimeOnly realValue;

            if (value is TimeOnly timeonly)
                realValue = timeonly;
            else
                throw new InvalidOperationException("Invalid value type");

            double totalSeconds = realValue.Ticks / 10000000;

            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(totalSeconds / 86400),
                CellReference = formatter.CellReference,
                StyleIndex = (uint)SpreadsheetFormat.TimeOnly
            };
        }

        private static Cell CreateCellDateOnly(object? value, SpreadsheetFormatter formatter)
        {
            DateOnly realValue;

            if (value is DateOnly dateonly)
                realValue = dateonly;
            else
                throw new InvalidOperationException("Invalid value type");

            return new Cell()
            {
                DataType = CellValues.Date,
                CellValue = new CellValue(realValue.ToDateTime(TimeOnly.MinValue)),
                CellReference = formatter.CellReference,
                StyleIndex = (uint)SpreadsheetFormat.DateOnly
            };
        }

#endif
    }
}
