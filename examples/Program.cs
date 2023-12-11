// See https://aka.ms/new-console-template for more information

using System;
using System.Threading.Tasks;
using BlackDigital.Report;
using System.Collections.Generic;
using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Example.Resources;
using System.Globalization;

using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using BlackDigital.Report.Example;

List<TestModel> list = new();

#if NET6_0_OR_GREATER

list.Add(new("Line <1> \r\n &A", 10, DateTime.Today, TimeSpan.FromHours(3), DateOnly.FromDateTime(DateTime.Today), new(3, 0)));
list.Add(new("Line <2> \r\n &B", -10, DateTime.Now, TimeSpan.FromMinutes(12), DateOnly.FromDateTime(DateTime.Today), new(5, 30)));
list.Add(new("Line <3> \n\r &C", 10.6d, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31)), DateOnly.FromDateTime(DateTime.Today), new(17, 45)));

List<string> headers = new()
{
    "Column 1",
    "Column 2",
    "Column 3",
    "Column 4",
    "Column 5",
    "Column 6",
    "Column 7",
};

#else

list.Add(new("Line 1", 10, DateTime.Today, TimeSpan.FromHours(3)));
list.Add(new("Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12)));
list.Add(new("Line 3", 10.6d, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31))));

List<string> headers = new()
{
    "Column 1",
    "Column 2",
    "Column 3",
    "Column 4"
};

#endif

var list2 = new List<List<object>>();

list2.Add(new List<object>() { "Line 1", 10, DateTime.Today, TimeSpan.FromHours(3) });
list2.Add(new List<object>() { "Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12) });
list2.Add(new List<object>() { "Line 3", 10.6m, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31)) });

List<string> headers2 = new()
{
    "Column 1",
    "Column 2",
    "Column 3",
    "Column 4"
};





System.ComponentModel.DataAnnotations.DisplayAttribute a;
//System.ComponentModel.DataAnnotations.DisplayColumnAttribute b;
//System.ComponentModel.DataAnnotations.DisplayFormatAttribute c;
//System.Text.Json.Serialization.JsonStringEnumConverter
//string a;

var task = ReportGenerator.Spreadsheet()
                .SetCompany("BlackDigital")
                .SetApplication("BlackDigital Report")
                .SetManager("TioRAC Lab")
                .SetAppVersion("1.1")

                .SetTitle("My Title")
                .SetSubject("My Subject")
                .AddCreators("TioRAC")
                .AddCreators("Other Creator")
                .AddKeyword("Spreadsheet")
                .AddKeyword("Generate")
                .SetDescription("A spreadshet test generator")
                .SetCategory("Report Category")

                .AddSheet("First")
                .AddTable("Data")
                .AddHeader(headers)
                .AddBody(list)
                .Workbook()
                .AddSheet("Second")
                .AddValue("My text header")
                .AddTable("Data2", "B3")
                .AddHeader(headers2)
                .AddBody(list2)
                .Sheet()
                .AddTable("Data3", "g4")
                .AddHeader(headers)
                .AddBody(list)
                .BuildAsync("D:\\Teste\\temp\\test-default.xlsx");

task.Wait();

return;

/*task = ReportGenerator.Spreadsheet()
                .SetCompany("BlackDigital")
                .SetResourceManager(Texts.ResourceManager)
                .SetFormatProvider(new CultureInfo("pt"))
                .AddSheet("First")
                .AddTable("Data")
                .FillObject(list)
                .Spreadsheet()
                .AddSheet("Second")
                .AddValue("My text header")
                .AddTable("Data2", "B3")
                .AddHeader(headers)
                .Fill(list2)
                .Sheet()
                .AddTable("Data3", "g4")
                .FillObject(list)
                .BuildAsync("test-pt.xlsx");*/

task.Wait();