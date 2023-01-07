using System;
using System.Diagnostics.CodeAnalysis;

namespace BlackDigital.Report.Spreadsheet
{
    public readonly struct SheetPosition : IComparable,
                                           IComparable<SheetPosition>,
                                           IEquatable<SheetPosition>
    {
        public SheetPosition(uint column, uint row)
        {
            if (column < 1)
                throw new ArgumentOutOfRangeException(nameof(column), "Column must be greater than 0.");

            if (row < 1)
                throw new ArgumentOutOfRangeException(nameof(row), "Row must be greater than 0.");

            Column = column;
            Row = row;
        }

        public readonly uint Column;

        public readonly uint Row;
        
        public static bool operator ==(SheetPosition left, SheetPosition right) => left.Column == right.Column && left.Row == right.Row;
        
        public static bool operator !=(SheetPosition left, SheetPosition right) => !(left == right);

        public override readonly bool Equals([NotNullWhen(true)] object? obj) => obj is SheetPosition && Equals((SheetPosition)obj);

        public readonly bool Equals(SheetPosition other) => this == other;

        public override readonly int GetHashCode() => HashCode.Combine(Column, Row);

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
    }
}
