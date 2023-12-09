using System;
using System.Linq;
using System.Diagnostics.CodeAnalysis;

namespace BlackDigital.Report.Spreadsheet
{
    public readonly struct SheetPosition : IComparable,
                                           IComparable<SheetPosition>,
                                           IEquatable<SheetPosition>
    {
        #region "Constructors"

        public SheetPosition(uint column, uint row)
        {
            if (column < 1)
                throw new ArgumentOutOfRangeException(nameof(column), "Column must be greater than 0.");

            if (row < 1)
                throw new ArgumentOutOfRangeException(nameof(row), "Row must be greater than 0.");

            Column = column;
            Row = row;
        }

        public SheetPosition(string cellReference)
        {
            var columnString = new String(cellReference.Where(Char.IsLetter).ToArray());
            var rowString = new String(cellReference.Where(Char.IsDigit).ToArray());

            if (!UInt32.TryParse(rowString, out uint row))
                throw new ArgumentException("Invalid cell reference");
            
            Column = LetterToNumber(columnString);
            Row = row;

            if (Column < 1)
                throw new ArgumentOutOfRangeException(nameof(Column), "Column must be greater than 0.");

            if (Row < 1)
                throw new ArgumentOutOfRangeException(nameof(Row), "Row must be greater than 0.");
        }

        #endregion

        #region Properties

        public readonly uint Column;

        public readonly uint Row;

        #endregion Properties

        #region "Methods"

        public SheetPosition AddColumn(uint value = 1)
        {
            return new SheetPosition(Column + value, Row);
        }

        public SheetPosition AddRow(uint value = 1)
        {
            return new SheetPosition(Column, Row + value);
        }

        public SheetPosition Add(uint column, uint row)
        {
            return new SheetPosition(Column + column, Row + row);
        }

        public uint CountColumns(SheetPosition position)
        {
            var max = Math.Max(Column, position.Column);
            var min = Math.Min(Column, position.Column);

            return max - min + 1;
        }

        public uint CountRows(SheetPosition position)
        {
            var max = Math.Max(Row, position.Row);
            var min = Math.Min(Row, position.Row);

            return max - min + 1;
        }

        private string GetLetterPosition()
        {
            string columnString = "";
            decimal columnNumber = Column;

            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }

        private static uint LetterToNumber(string column)
        {
            uint retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                uint colNum = (uint)(colPiece - 64);
                retVal += colNum * (uint)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        #endregion "Methods"

        #region "Operators"
        
        public int CompareTo(object? obj)
        {
            if (obj == null)
                return 1;

            if (obj is SheetPosition other)
                return CompareTo(other);

            throw new ArgumentException("Object is not a SheetPosition.");
        }

        public int CompareTo(SheetPosition other)
        {
            if (Column < other.Column)
                return -1;

            if (Column > other.Column)
                return 1;

            if (Column == other.Column)
            {
                if (Row < other.Row)
                    return -1;

                if (Row > other.Row)
                    return 1;
            }

            return 0;
        }

        public override readonly bool Equals([NotNullWhen(true)] object? obj) => obj is SheetPosition && Equals((SheetPosition)obj);

        public readonly bool Equals(SheetPosition other) => this == other;

        public override string ToString()
        {
            return $"{GetLetterPosition()}{Row}";
        }

        public override readonly int GetHashCode() => HashCode.Combine(Column, Row);

        public static bool operator ==(SheetPosition left, SheetPosition right) 
            => left.Column == right.Column && left.Row == right.Row;
        
        public static bool operator !=(SheetPosition left, SheetPosition right) 
            => !(left == right);

        public static bool operator <(SheetPosition left, SheetPosition right) 
            => left.CompareTo(right) < 0;

        public static bool operator >(SheetPosition left, SheetPosition right) 
            => left.CompareTo(right) > 0;

        public static bool operator <=(SheetPosition left, SheetPosition right) 
            => left.CompareTo(right) <= 0;

        public static bool operator >=(SheetPosition left, SheetPosition right) 
            => left.CompareTo(right) >= 0;

        public static implicit operator string(SheetPosition sheetPosition)
        {
            return sheetPosition.ToString();
        }

        public static implicit operator SheetPosition(string cellReference)
        {
            return new SheetPosition(cellReference);
        }

        #endregion "Operators"
    }
}
