// See https://aka.ms/new-console-template for more information

using BlackDigital.Report;
using DocumentFormat.OpenXml;

var list = new List<List<object>>();

list.Add(new List<object>() { "Line 1", 10, DateTime.Today, TimeSpan.FromHours(3) });
list.Add(new List<object>() { "Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12) });
list.Add(new List<object>() { "Line 3", 10.6m, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31)) });


var report = Report.Spreadsheet().SetCompany("BlackDigital")
                    .SetType(SpreadsheetDocumentType.Workbook)
                    .GenerateAsync(list, new List<string> { "Column A", "Column B", "Column C", "Column D" });
    
File.WriteAllBytes(@"d:\teste\test.xlsx", report.Result);