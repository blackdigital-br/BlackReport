using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Spreadsheet
{
    public class SpreadsheetReport : IReport
    {
        public Task<byte[]> GenerateReportAsync<T>(IEnumerable<T> data, ReportArguments arguments)
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> GenerateReportAsync(IEnumerable<IEnumerable<object>> data, IEnumerable<string> columns, ReportArguments arguments)
        {
            return Task.FromResult(CreateSpreadsheet(data, columns, (SpreadsheetArguments)arguments));
        }

        

        private byte[] CreateSpreadsheet(IEnumerable<IEnumerable<object>> data, 
                                       IEnumerable<string> columns, 
                                       SpreadsheetArguments arguments)
        {
            using MemoryStream memoryStream = new();
            using SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, arguments.DocumentType);
            CreateParts(data, columns, arguments, document);

            document.Close();

            return memoryStream.ToArray();
        }

        private void CreateParts(IEnumerable<IEnumerable<object>> data,
                                       IEnumerable<string> columns,
                                       SpreadsheetArguments arguments,
                                       SpreadsheetDocument document)
        {
            HashSet<string> sharedStrings = new();

            FilePropertiesPart(document, arguments);
            //CoreFilePropertiesPart(document, arguments);

            WorkbookPart workbookPart = WorkbookPart(data, columns, arguments, document);
            WorksheetPart(data, columns, arguments, document, workbookPart, sharedStrings);

            workbookPart.Workbook.Save();
        }

        private void FilePropertiesPart(SpreadsheetDocument document, SpreadsheetArguments arguments)
        {
            var filePropertiesPart = document.AddExtendedFilePropertiesPart();
            Properties properties = new();

            if (!string.IsNullOrWhiteSpace(arguments.Company))
            {
                Company company = new(arguments.Company);
                properties.Append(company);
            }

            filePropertiesPart.Properties = properties;
        }

        private void CoreFilePropertiesPart(SpreadsheetDocument document, SpreadsheetArguments arguments)
        {
            var coreFilePropertiesPart = document.AddCoreFilePropertiesPart();
        }

        private WorkbookPart WorkbookPart(IEnumerable<IEnumerable<object>> data,
                                       IEnumerable<string> columns,
                                       SpreadsheetArguments arguments,
                                       SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            GenerateWorkbookStylesPart(workbookPart);

            return workbookPart;
        }

        private void WorksheetPart(IEnumerable<IEnumerable<object>> data,
                                       IEnumerable<string> columns,
                                       SpreadsheetArguments arguments,
                                       SpreadsheetDocument document,
                                       WorkbookPart workbookPart,
                                       HashSet<string> sharedStrings)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new();
            Worksheet worksheet = new(sheetData);
            worksheetPart.Worksheet = worksheet;

            Sheets xSheets = workbookPart.Workbook.AppendChild<Sheets>(new());
            Sheet xSheet = new()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet Name"
            };


            CreateWorksheetColumns(columns, worksheet, sheetData);
            CreateWorkSheetDataRows(data, sheetData);


            


            xSheets.Append(xSheet);
            

            /*SheetDimension xSheetDimension = new();
            xSheetDimension.Reference = "A1:K4";
            xWorksheet.Append(xSheetDimension);*/

            /*SheetViews xSheetViews = new();
            SheetView xSheetView = new();
            xSheetView.TabSelected = true;
            xSheetView.WorkbookViewId = 0u;
            xSheetViews.Append(xSheetView);
            xWorksheet.Append(xSheetViews);*/
        }

        private void GenerateWorkbookStylesPart(WorkbookPart workbookPart)
        {
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = new();


            NumberingFormats numberingFormats = new();
            numberingFormats.Count = 2u;

            NumberingFormat numberingFormat = new();
            numberingFormat.NumberFormatId = 164u;
            numberingFormat.FormatCode = "[$-F400]h:mm:ss\\ AM/PM";

            numberingFormats.Append(numberingFormat);

            numberingFormat = new NumberingFormat();
            numberingFormat.NumberFormatId = 165u;
            numberingFormat.FormatCode = "d/m/yy\\ h:mm;@";

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

            CellFormats xCellFormats = new CellFormats();
            xCellFormats.Count = 3u;

            xCellFormat = new CellFormat();
            xCellFormat.NumberFormatId = 0u;
            xCellFormat.FontId = 0u;
            xCellFormat.FillId = 0u;
            xCellFormat.BorderId = 0u;
            xCellFormat.FormatId = 0u;

            xCellFormats.Append(xCellFormat);

            xCellFormat = new CellFormat();
            xCellFormat.NumberFormatId = 164u;
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

            stylesheet.Append(xCellFormats);

            CellStyles xCellStyles = new CellStyles();
            xCellStyles.Count = 1u;

            CellStyle xCellStyle = new CellStyle();
            xCellStyle.Name = "Normal";
            xCellStyle.FormatId = 0u;
            xCellStyle.BuiltinId = 0u;

            xCellStyles.Append(xCellStyle);

            stylesheet.Append(xCellStyles);

            DifferentialFormats xDifferentialFormats = new DifferentialFormats();
            xDifferentialFormats.Count = 3u;

            DifferentialFormat xDifferentialFormat = new DifferentialFormat();

            numberingFormat = new NumberingFormat();
            numberingFormat.NumberFormatId = 164u;
            numberingFormat.FormatCode = "[$-F400]h:mm:ss\\ AM/PM";

            xDifferentialFormat.Append(numberingFormat);

            xDifferentialFormats.Append(xDifferentialFormat);

            xDifferentialFormat = new DifferentialFormat();

            numberingFormat = new NumberingFormat();
            numberingFormat.NumberFormatId = 165u;
            numberingFormat.FormatCode = "d/m/yy\\ h:mm;@";

            xDifferentialFormat.Append(numberingFormat);

            xDifferentialFormats.Append(xDifferentialFormat);

            xDifferentialFormat = new DifferentialFormat();

            numberingFormat = new NumberingFormat();
            numberingFormat.NumberFormatId = 165u;
            numberingFormat.FormatCode = "d/m/yy\\ h:mm;@";

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

        private void CreateWorksheetColumns(IEnumerable<string> columns, Worksheet xWorksheet, SheetData xSheetData)
        {
            /*Columns xColumns = new();

            foreach (string column in columns)
            {
                Column xColumn = new();
                xColumn.Width = 15;
                xColumn.Min = 1;
                xColumn.Max = 1;
                xColumn.CustomWidth = true;
                xColumns.Append(xColumn);
            }

            xWorksheet.Append(xColumns);*/
            CreateWorksheetRow(columns, xSheetData, 1);
        }

        private void CreateWorkSheetDataRows(IEnumerable<IEnumerable<object>> data, SheetData xSheetData)
        {
            uint line = 2;
            foreach (var item in data)
            {
                CreateWorksheetRow(item, xSheetData, line++);
            }
        }

        private void CreateWorksheetRow(IEnumerable<object> row, SheetData xSheetData, uint line)
        {
            Row xRow = new();
            xRow.RowIndex = line;
            var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var letterPos = 0;

            foreach (object cell in row)
            {
                var letter = letters[letterPos++];

                xRow.Append(CreateCellByType(cell, $"{letter}{line}"));
            }

            xSheetData.Append(xRow);
        }

        private Cell CreateCellByType(object value, string cellReference)
        {
            if (value is string)
                return CreateCellString(value.ToString(), cellReference);
            else if (value is DateTime)
                return CreateCellDateTime((DateTime)value, cellReference);
            else if (value is DateTimeOffset)
                return CreateCellDateTime(((DateTimeOffset)value).DateTime, cellReference);
            else if (value is TimeSpan)
                return CreateCellTimespan(((TimeSpan)value), cellReference);
            else if (value is int || value is short)
                return CreateCellInt(Convert.ToInt32(value), cellReference);
            else if (value is double || value is decimal || value is float)
                return CreateCellDouble(Convert.ToDouble(value), cellReference);

            return CreateCellNull(cellReference);
        }

        private Cell CreateCellNull(string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(""),
                CellReference = cellReference
            };
        }

        private Cell CreateCellString(string value, string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(value),
                CellReference = cellReference
            };
        }

        private Cell CreateCellDateTime(DateTime value, string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.Date,
                CellValue = new CellValue(value),
                CellReference = cellReference,
                StyleIndex = 2u
            };
        }

        private Cell CreateCellTimespan(TimeSpan value, string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.Number,
                //CellValue = new CellValue((new DateTime(1900, 1, 1)).Add(value)),
                CellValue = new CellValue(value.TotalSeconds / 86400),
                CellReference = cellReference,
                StyleIndex = 1u
            };
        }

        private Cell CreateCellInt(int value, string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(value),
                CellReference = cellReference
            };
        }

        private Cell CreateCellDouble(double value, string cellReference)
        {
            return new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(value),
                CellReference = cellReference
            };
        }
    }
}
