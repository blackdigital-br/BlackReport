using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    internal static class SpreadsheetHelper
    {
        internal static string NumberToExcelColumn(uint column)
        {
            string columnString = "";
            decimal columnNumber = column;
            
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            
            return columnString;
        }

        internal static uint ExcelColumnToNumber(string column)
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

        internal static (uint column, uint row) CellReferenceToNumbers(string cellReference)
        {
            var columnString = new String(cellReference.Where(Char.IsLetter).ToArray());
            var rowString = new String(cellReference.Where(Char.IsDigit).ToArray());

            if (!UInt32.TryParse(rowString, out uint row))
                throw new ArgumentException("Invalid cell reference");

            uint column = SpreadsheetHelper.ExcelColumnToNumber(columnString);

            return (column, row);
        }
    }
}
