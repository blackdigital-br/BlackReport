// See https://aka.ms/new-console-template for more information

using BlackDigital.Report;
using BlackDigital.Report.Example.Model;
using DocumentFormat.OpenXml;

List<TestModel> list = new();
list.Add(new("Line 1", 10, DateTime.Today, TimeSpan.FromHours(3)));
list.Add(new("Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12)));
list.Add(new("Line 3", 10.6d, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31))));

var list2 = new List<List<object>>();

list2.Add(new List<object>() { "Line 1", 10, DateTime.Today, TimeSpan.FromHours(3) });
list2.Add(new List<object>() { "Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12) });
list2.Add(new List<object>() { "Line 3", 10.6m, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31)) });

List<string> headers = new()
{
    "Column 1",
    "Column 2",
    "Column 3",
    "Column 4"
};



var report = ReportGenerator.Spreadsheet()
                            .SetCompany("BlackDigital")
                            .SetType(SpreadsheetDocumentType.Workbook)
                            .AddSheet("First")
                            .AddTable("Data")
                            .FillObject(list)
                            .Spreadsheet()
                            .AddSheet("Second")
                            .AddValue("My text header")
                            .AddTable("Data2", "B3")
                            .AddHeader(headers)
                            .Fill(list2)
                            .BuildAsync();
    
File.WriteAllBytes(@"d:\teste\test.xlsx", report.Result);