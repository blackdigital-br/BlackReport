# BlackDigital Report

Simple report generator using OpenXML.

Currently it is only generating reports in spreadsheets.

**Alpha version, can be changed at any time without notice.**

```csharp
    using BlackDigital.Report;

    ReportGenerator.Spreadsheet()
                   .AddSheet("First")
                   .AddTable("Data")
                   .FillObject(list)
                   .BuildAsync(@"test.xlsx");
```

![Example](https://raw.githubusercontent.com/blackdigital-br/BlackReport/main/docs/images/ClassExample.png)

See more on the [Wiki](https://github.com/blackdigital-br/BlackReport/wiki)

* [Installing](#installing)
* [Coding](#coding)
    * [From class list](#from-class-list)
    * [From List](#from-list)
    * [Others Examples](#others-examples)
* [Roadmap](#roadmap)


## Installing

* Package Manager

```
Install-Package BlackDigital.Report
```

* .NET CLI

```
dotnet add package BlackDigital.Report
```

## Coding

### From class list

1. Class example:

```csharp
    public class TestModel
    {
        public TestModel(string name, double number, DateTime objDate, TimeSpan time)
        {
            Name = name;
            Number = number;
            ObjDate = objDate;
            Time = time;            
        }

        public string Name { get; set; }
        public double Number { get; set; }
        public DateTime ObjDate { get; set; }
        public TimeSpan Time { get; set; }
    }
```

2. Create class instance list:

```csharp
    List<TestModel> list = new();
    list.Add(new("Line 1", 10, DateTime.Today, TimeSpan.FromHours(3)));
    list.Add(new("Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12)));
    list.Add(new("Line 3", 10.6d, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31))));
```

3. Generate spreadsheets report:

```csharp
    using BlackDigital.Report;

    ReportGenerator.Spreadsheet()
                   .AddSheet("First")
                   .AddTable("Data")
                   .FillObject(list)
                   .BuildAsync(@"test.xlsx");
```

![Example](https://raw.githubusercontent.com/blackdigital-br/BlackReport/main/docs/images/ClassExample.png)

### From List

1. Create object in two-dimensional array:

```csharp
    var list = new List<List<object>>();

    list.Add(new List<object>() { "Line 1", 10, DateTime.Today, TimeSpan.FromHours(3) });
    list.Add(new List<object>() { "Line 2", -10, DateTime.Now, TimeSpan.FromMinutes(12) });
    list.Add(new List<object>() { "Line 3", 10.6m, DateTime.UtcNow, TimeSpan.FromMinutes(45).Add(TimeSpan.FromSeconds(31)) });
```

2. Create headers list:

```csharp
    List<string> headers = new()
    {
        "Column 1",
        "Column 2",
        "Column 3",
        "Column 4"
    };
```

3. Generate spreadsheets report:

```csharp
    using BlackDigital.Report;

    ReportGenerator.Spreadsheet()
                   .AddSheet("Second")
                   .AddTable("Data2", "B3")
                   .AddHeader(headers)
                   .Fill(list) 
                   .BuildAsync(@"test.xlsx");
```

![Example](https://raw.githubusercontent.com/blackdigital-br/BlackReport/main/docs/images/ListExample.png)

### Others Examples

#### Add single value:

```csharp
    using BlackDigital.Report;

    ReportGenerator.Spreadsheet()
                   .AddSheet("My sheet")
                   .AddValue("My text header")
                   .AddValue("Column B Line 5", "B5")
                   .AddValue("Column C Line 5", 3, 5);
```

#### Builders Examples

```csharp
    using BlackDigital.Report;

    ReportGenerator.Spreadsheet()
                   ...
                   .BuildAsync(@"test.xlsx"); //Save in file

    MemoryStream ms = new();
    ReportGenerator.Spreadsheet()
                   ...
                   .BuildAsync(ms); //Save in stream

    byte[] buffer = await ReportGenerator.Spreadsheet()
                            ...
                            .BuildAsync(); //Return as byte array
    
```

### Example

https://github.com/blackdigital-br/BlackReport/tree/main/src/BlackDigital.Report.Example

## Roadmap

    ☑️ Excel Tables. (0.1.0)
    ☑️ Fill from instance class. (0.1.0)
    ☑️ Multiple tables in the same worksheet. (0.1.1)
    ☑️ .NET5 (0.1.2)
	☑️ Header Globalization. (0.1.2)
    ◼️ Use DisplayAttribute to get name of columns and properties that should be generated.
    ◼️ Excel Shared String.
    ◼️ Cells with formulas.
    ◼️ Cell value event.
    ◼️ Tables footers.
    ◼️ Unit test.
    ◼️ Olders .net versions.
    ◼️ Others types (Word, csv...).
    ◼️ Your suggestion.


Everyone is welcome to help or report bugs. 💪    