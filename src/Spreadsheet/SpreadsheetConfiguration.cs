using System;
using System.Collections.Generic;
using BlackDigital.Report.Spreadsheet.Formatter;


namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetConfiguration
    {
        #region "Constructor"

        public SpreadsheetConfiguration()
        {
            SharedStringsSize = 64;
            CreateBaseValues();
        }

        #endregion "Constructor"

        #region "Properties"

        private Dictionary<Type, ICreaterCellValue> CreaterValues { get; } = new();

        internal uint? SharedStringsSize { get; private set; }

        #endregion "Properties"

        #region "Configuration"

        public SpreadsheetConfiguration AddCreaterCellValue(Type type, ICreaterCellValue createrCellValue)
        {
            if (CreaterValues.ContainsKey(type))
                CreaterValues[type] = createrCellValue;
            else
                CreaterValues.Add(type, createrCellValue);

            return this;
        }

        public SpreadsheetConfiguration AddCreaterCellValue<T>(ICreaterCellValue createrCellValue)
            => AddCreaterCellValue(typeof(T), createrCellValue);

        /// <summary>
        /// Set string limit size for shared strings
        /// </summary>
        /// <param name="size">Size of strings, 0 not use shared strings, null not limit</param>
        /// <returns></returns>
        public SpreadsheetConfiguration SetSharedStringsSize(uint? size)
        {
            SharedStringsSize = size;
            return this;
        }

        #endregion "Configuration"

        #region "Methods"

        private void CreateBaseValues()
        {
            CreaterValues.Add(typeof(bool), new BooleanCreaterCellValue());
            CreaterValues.Add(typeof(char), new StringCreaterCellValue(this));
            CreaterValues.Add(typeof(string), new StringCreaterCellValue(this));
            CreaterValues.Add(typeof(byte), new NumberCreateCellValue());
            CreaterValues.Add(typeof(sbyte), new NumberCreateCellValue());
            CreaterValues.Add(typeof(ushort), new NumberCreateCellValue());
            CreaterValues.Add(typeof(short), new NumberCreateCellValue());
            CreaterValues.Add(typeof(int), new NumberCreateCellValue());
            CreaterValues.Add(typeof(uint), new NumberCreateCellValue());
            CreaterValues.Add(typeof(ulong), new NumberCreateCellValue());
            CreaterValues.Add(typeof(long), new NumberCreateCellValue());
            CreaterValues.Add(typeof(float), new NumberCreateCellValue());
            CreaterValues.Add(typeof(double), new NumberCreateCellValue());
            CreaterValues.Add(typeof(decimal), new NumberCreateCellValue());
            CreaterValues.Add(typeof(DateTime), new DateTimeCreaterCellValue());
            CreaterValues.Add(typeof(DateTimeOffset), new DateTimeCreaterCellValue());
            CreaterValues.Add(typeof(TimeSpan), new TimeSpanCreaterCellValue());

#if NET6_0_OR_GREATER
            CreaterValues.Add(typeof(TimeOnly), new TimeOnlyCreaterCellValue());
            CreaterValues.Add(typeof(DateOnly), new DateOnlyCreaterCellValue());
#endif
        }

        public ICreaterCellValue GetCreaterCellValue(Type type)
        {
            if (CreaterValues.ContainsKey(type))
                return CreaterValues[type];

            return new StringCreaterCellValue(this);
        }

        #endregion "Methods"
    }
}
