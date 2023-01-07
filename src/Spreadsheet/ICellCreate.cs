using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public interface ICellCreate
    {
        Cell CreateCell(object? value, SpreadsheetFormatter formatter);
    }
}
