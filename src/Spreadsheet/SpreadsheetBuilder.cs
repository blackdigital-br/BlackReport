using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Resources;
using System.Globalization;
using System.Threading.Tasks;
using System.IO.Packaging;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetBuilder : ReportBuilder
    {
        #region "Constructor"
        
        public SpreadsheetBuilder() { } //: base(new SpreadsheetGenerator()) { }

        #endregion "Constructor"

        #region "Properties"

        private SpreadsheetDocumentType DocumentType { get; set; } = SpreadsheetDocumentType.Workbook;

        private string Company { get; set; } = null;

        internal List<SheetBuilder> Sheets { get; private set; } = new();

        internal List<TableBuilder> Tables { get; private set; } = new();

        #endregion "Properties"

        #region "Builder"

        public SpreadsheetBuilder SetResourceManager(ResourceManager resource) 
            => this.SetResourceManager<SpreadsheetBuilder>(resource);

        public SpreadsheetBuilder SetFormatProvider(IFormatProvider formatProvider)
            => this.SetFormatProvider<SpreadsheetBuilder>(formatProvider);

        public SpreadsheetBuilder SetCompany(string company)
        {
            Company = company;
            return this;
        }

        public SpreadsheetBuilder SetType(SpreadsheetDocumentType type)
        {
            DocumentType = type;
            return this;
        }

        public SheetBuilder AddSheet(string name)
        {
            var sheet = new SheetBuilder(this, name);
            Sheets.Add(sheet);
            return sheet;
        }

        #endregion "Builder"

        #region "Generate"

        public override Task<byte[]> BuildAsync()
        {
            return Task.Run(() => {
                using MemoryStream memoryStream = new();
                CreateSpreadsheet(memoryStream);
                return memoryStream.ToArray();
            });
        }

        public override Task BuildAsync(string file)
        {
            return Task.Run(() => {
                using var fs = File.Create(file);
                CreateSpreadsheet(fs);
            });
        }

        private void CreateSpreadsheet(Stream stream)
        {
            using var package = Package.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using SpreadsheetDocument document = SpreadsheetDocument.Create(package, DocumentType);
            CreateParts(document);
            document.Close();
        }


        /*private void CreateEmptyFile(string file)
        {
            using var fs = File.Create(file);
            using var package = Package.Open(fs, FileMode.Create, FileAccess.Write);
            using var excel = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook))
        }*/

        private void CreateParts(SpreadsheetDocument document)
        {
            HashSet<string> sharedStrings = new();

            FilePropertiesPart(document);
            //CoreFilePropertiesPart(document, arguments);
            
            WorkbookPart workbookPart = WorkbookPart(document);
            GenerateWorkbookStylesPart(workbookPart);
            Sheets sheets = workbookPart.Workbook.AppendChild<Sheets>(new());

            foreach (var sheet in Sheets)
            {
                sheet.GenerateWorksheetPart(workbookPart, sheets, sharedStrings);
            }

            workbookPart.Workbook.Save();
        }

        private void FilePropertiesPart(SpreadsheetDocument document)
        {
            var filePropertiesPart = document.AddExtendedFilePropertiesPart();
            Properties properties = new();

            if (!string.IsNullOrWhiteSpace(Company))
            {
                Company company = new(Company);
                properties.Append(company);
            }

            filePropertiesPart.Properties = properties;
        }

        private void CoreFilePropertiesPart(SpreadsheetDocument document)
        {
            var coreFilePropertiesPart = document.AddCoreFilePropertiesPart();
        }

        private WorkbookPart WorkbookPart(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            return workbookPart;
        }

        private void GenerateWorkbookStylesPart(WorkbookPart workbookPart)
        {
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = new();


            NumberingFormats numberingFormats = new();
            numberingFormats.Count = 3u;

            NumberingFormat numberingFormat = new();
            numberingFormat.NumberFormatId = 168u;
            numberingFormat.FormatCode = "[h]:mm:ss;@";
            numberingFormats.Append(numberingFormat);

            numberingFormat = new();
            numberingFormat.NumberFormatId = 165u;
            numberingFormat.FormatCode = "d/m/yy\\ h:mm;@";
            numberingFormats.Append(numberingFormat);

            numberingFormat = new();
            numberingFormat.NumberFormatId = 164u;
            numberingFormat.FormatCode = "[$-F400]h:mm:ss\\ AM/PM";
            numberingFormats.Append(numberingFormat);

            stylesheet.Append(numberingFormats);


            
            Fonts xFonts = new Fonts();
            xFonts.Count = 1u;
            xFonts.KnownFonts = true;

            Font xFont = new Font();

            FontSize xFontSize = new FontSize();
            xFontSize.Val = 11D;

            xFont.Append(xFontSize);

            Color xColor = new Color();
            xColor.Theme = 1u;

            xFont.Append(xColor);

            FontName xFontName = new FontName();
            xFontName.Val = "Calibri";

            xFont.Append(xFontName);

            FontFamilyNumbering xFontFamilyNumbering = new FontFamilyNumbering();
            xFontFamilyNumbering.Val = 2;

            xFont.Append(xFontFamilyNumbering);

            FontScheme xFontScheme = new FontScheme();
            xFontScheme.Val = FontSchemeValues.Minor;

            xFont.Append(xFontScheme);

            xFonts.Append(xFont);

            stylesheet.Append(xFonts);

            Fills xFills = new Fills();
            xFills.Count = 2u;

            Fill xFill = new Fill();

            PatternFill xPatternFill = new PatternFill();
            xPatternFill.PatternType = PatternValues.None;

            xFill.Append(xPatternFill);

            xFills.Append(xFill);

            xFill = new Fill();

            xPatternFill = new PatternFill();
            xPatternFill.PatternType = PatternValues.Gray125;

            xFill.Append(xPatternFill);

            xFills.Append(xFill);

            stylesheet.Append(xFills);

            Borders xBorders = new Borders();
            xBorders.Count = 1u;

            Border xBorder = new Border();

            LeftBorder xLeftBorder = new LeftBorder();

            xBorder.Append(xLeftBorder);

            RightBorder xRightBorder = new RightBorder();

            xBorder.Append(xRightBorder);

            TopBorder xTopBorder = new TopBorder();

            xBorder.Append(xTopBorder);

            BottomBorder xBottomBorder = new BottomBorder();

            xBorder.Append(xBottomBorder);

            DiagonalBorder xDiagonalBorder = new DiagonalBorder();

            xBorder.Append(xDiagonalBorder);

            xBorders.Append(xBorder);

            stylesheet.Append(xBorders);

            CellStyleFormats xCellStyleFormats = new CellStyleFormats();
            xCellStyleFormats.Count = 1u;

            CellFormat xCellFormat = new CellFormat();
            xCellFormat.NumberFormatId = 0u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;

            xCellStyleFormats.Append(xCellFormat);

            stylesheet.Append(xCellStyleFormats);

            CellFormats xCellFormats = new();
            xCellFormats.Count = 5u;

            xCellFormat = new();
            xCellFormat.NumberFormatId = 0u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;
            xCellFormats.Append(xCellFormat);

            xCellFormat = new();
            xCellFormat.NumberFormatId = 14u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;
            xCellFormat.ApplyNumberFormat = true;
            xCellFormats.Append(xCellFormat);

            xCellFormat = new();
            xCellFormat.NumberFormatId = 168u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;
            xCellFormat.ApplyNumberFormat = true;
            xCellFormats.Append(xCellFormat);

            xCellFormat = new CellFormat();
            xCellFormat.NumberFormatId = 165u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;
            xCellFormat.ApplyNumberFormat = true;
            xCellFormats.Append(xCellFormat);

            xCellFormat = new();
            xCellFormat.NumberFormatId = 164u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;
            xCellFormat.ApplyNumberFormat = true;
            xCellFormats.Append(xCellFormat);

            stylesheet.Append(xCellFormats);

            CellStyles xCellStyles = new CellStyles();
            xCellStyles.Count = 1u;

            CellStyle xCellStyle = new CellStyle();
            xCellStyle.Name = "Normal";
            xCellStyle.FormatId = 0u;
            xCellStyle.BuiltinId = 0u;

            xCellStyles.Append(xCellStyle);

            stylesheet.Append(xCellStyles);



            
            DifferentialFormats xDifferentialFormats = new();
            xDifferentialFormats.Count = 3u;

            
            DifferentialFormat xDifferentialFormat = new();
            numberingFormat = new();
            numberingFormat.NumberFormatId = 168u;
            numberingFormat.FormatCode = "[h]:mm:ss;@";
            xDifferentialFormat.Append(numberingFormat);
            xDifferentialFormats.Append(xDifferentialFormat);

            xDifferentialFormat = new();
            numberingFormat = new();
            numberingFormat.NumberFormatId = 165u;
            numberingFormat.FormatCode = "d/m/yy\\ h:mm;@";
            xDifferentialFormat.Append(numberingFormat);
            xDifferentialFormats.Append(xDifferentialFormat);

            xDifferentialFormat = new();
            numberingFormat = new();
            numberingFormat.NumberFormatId = 164u;
            numberingFormat.FormatCode = "[$-F400]h:mm:ss\\ AM/PM";
            xDifferentialFormat.Append(numberingFormat);
            xDifferentialFormats.Append(xDifferentialFormat);


            stylesheet.Append(xDifferentialFormats);






            
            TableStyles xTableStyles = new TableStyles();
            xTableStyles.Count = 0u;
            xTableStyles.DefaultTableStyle = "TableStyleMedium2";
            xTableStyles.DefaultPivotStyle = "PivotStyleLight16";

            stylesheet.Append(xTableStyles);

            workbookStylesPart.Stylesheet = stylesheet;
        }

        #endregion "Generate"
    }
}
