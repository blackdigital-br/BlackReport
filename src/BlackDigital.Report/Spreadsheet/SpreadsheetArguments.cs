using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetArguments : ReportArguments
    {
        public SpreadsheetArguments() : base(new SpreadsheetReport()) { }

        internal SpreadsheetDocumentType DocumentType { get; private set; } = SpreadsheetDocumentType.Workbook;

        internal string Company { get; private set; } = null;


        public SpreadsheetArguments SetCompany(string company)
        {
            Company = company;
            return this;
        }

        public SpreadsheetArguments SetType(SpreadsheetDocumentType type)
        {
            DocumentType = type;
            return this;
        }

    }
}
