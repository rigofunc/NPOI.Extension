Using `Fluent API` to configure POCO excel behaviors, and then provides IEnumerable&lt;T&gt; has save to and load from excel functionalities.

# Features
- [x] Decouple the configuration from the POCO model by using `fluent api`.
- [x] Support none configuration POCO, so that if English is your mother language, none any more configurations;

The first features will be very useful for English not their mother language developers.

# IMPORTAMT
1. The repo fork from my [NPOI.Extension](https://github.com/xyting/NPOI.Extension), and remove all the attributes based features (but can be extended, see following demo), and will only support `Fluent API`.
2. All the issues found in [NPOI.Extension](https://github.com/xyting/NPOI.Extension) will be and only be fixed by [FluentExcel](https://github.com/Arch/FluentExcel), so, please update your codes use `FluentExcel`.

# Overview

![FluentExcel demo](images/demo.PNG)

# Get Started

The following demo codes come from [sample](samples), download and run it for more information.

## Install FluentExcel package

        PM> Install-Package FluentExcel
    
## Using FluentExcel in code

        using FluentExcel;

## Saving IEnumerable&lt;T&gt; to excel.

```csharp
var excelFile = @"/Users/rigofunc/Documents/sample.xlsx";

// save to excel file
reports.ToExcel(excelFile);
```

## Loading IEnumerable&lt;T&gt; from excel.

```csharp
// load from excel
var loadFromExcel = Excel.Load<Report>(excelFile);       
```

## From Annotations by extenstion methods.

```csharp
var fluentConfiguration = Excel.Setting.For<Report>().FromAnnotations();
```

The following demo show how to extend the exist functionalities by extension methods. `NOTE:` the initial idea from @tupunco.

### Applying the annotations to the model

```csharp
public class Report
{
    [Display(Name = "城市")]
    public string City { get; set; }
    [Display(Name = "楼盘")]
    public string Building { get; set; }
    [Display(Name = "区域")]
    public string Area { get; set; }
    [Display(Name = "成交时间")]
    public DateTime HandleTime { get; set; }
    [Display(Name = "经纪人")]
    public string Broker { get; set; }
    [Display(Name = "客户")]
    public string Customer { get; set; }
    [Display(Name = "房源")]
    public string Room { get; set; }
    [Display(Name = "佣金(元)")]
    public decimal Brokerage { get; set; }
    [Display(Name = "收益(元)")]
    public decimal Profits { get; set; }
}
```

### Defines the extension methods.

```csharp
public static class FluentConfigurationExtensions
{
    public static FluentConfiguration<TModel> FromAnnotations<TModel>(this FluentConfiguration<TModel> fluentConfiguration) where TModel : class
    {
        var properties = typeof(TModel).GetProperties();
        foreach (var property in properties)
        {
            var pc = fluentConfiguration.Property(property);

            var display = property.GetCustomAttribute<DisplayAttribute>();
            if (display != null)
            {
                pc.HasExcelTitle(display.Name);
                if (display.GetOrder().HasValue)
                {
                    pc.HasExcelIndex(display.Order);
                }
            }
            else
            {
                pc.HasExcelTitle(property.Name);
            }

            var format = property.GetCustomAttribute<DisplayFormatAttribute>();
            if (format != null)
            {
                pc.HasDataFormatter(format.DataFormatString
                                          .Replace("{0:", "")
                                          .Replace("}", ""));
            }

            if (pc.Index < 0)
            {
                pc.HasAutoIndex();
            }
        }

        return fluentConfiguration;
    }
}
```
    
## Use Fluent Api to configure POCO's excel behaviors

We can use `fluent api` to configure the model excel behaviors.

```csharp
/// <summary>
/// Use fluent configuration api. (doesn't poison your POCO)
/// </summary>
static void FluentConfiguration()
{
    var fc = Excel.Setting.For<Report>();

    fc.HasStatistics("合计", "SUM", 6, 7)
      .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
      .HasFreeze(columnSplit: 2, rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

    fc.Property(r => r.City)
      .HasExcelIndex(0)
      .HasExcelTitle("城市")
      .IsMergeEnabled();

    // or
    //fc.Property(r => r.City).HasExcelCell(0,"城市", allowMerge: true);

    fc.Property(r => r.Building)
      .HasExcelIndex(1)
      .HasExcelTitle("楼盘")
      .IsMergeEnabled();

    // configures the ignore when exporting or importing.
    fc.Property(r => r.Area)
      .HasExcelIndex(8)
      .HasExcelTitle("Area")
      .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

    // or
    //fc.Property(r => r.Area).IsIgnored(8, "Area", formatter: null, exportingIsIgnored: false, importingIsIgnored: true);

    fc.Property(r => r.HandleTime)
      .HasExcelIndex(2)
      .HasExcelTitle("成交时间")
      .HasDataFormatter("yyyy-MM-dd");

    // or 
    //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", formatter: "yyyy-MM-dd", allowMerge: false);
    // or
    //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", "yyyy-MM-dd");


    fc.Property(r => r.Broker)
      .HasExcelIndex(3)
      .HasExcelTitle("经纪人");

    fc.Property(r => r.Customer)
      .HasExcelIndex(4)
      .HasExcelTitle("客户");

    fc.Property(r => r.Room)
      .HasExcelIndex(5)
      .HasExcelTitle("房源");

    fc.Property(r => r.Brokerage)
      .HasExcelIndex(6)
      .HasDataFormatter("￥0.00")
      .HasExcelTitle("佣金(元)");

    fc.Property(r => r.Profits)
      .HasExcelIndex(7)
      .HasExcelTitle("收益(元)");
}
```

```csharp
class Program
{
    static void Main(string[] args)
    {
        // global call this
        FluentConfiguration();

        // demo the extension point
        //var fc = Excel.Setting.For<Report>().FromAnnotations();

        var len = 20;
        var reports = new Report[len];
        for (int i = 0; i < len; i++)
        {
            reports[i] = new Report
            {
                City = "ningbo",
                Building = "世茂首府",
                HandleTime = DateTime.Now,
                Broker = "rigofunc 18957139**7",
                Customer = "yingting 18957139**7",
                Room = "2#1703",
                Brokerage = 125 * i,
                Profits = 25 * i
            };
        }

        var excelFile = @"/Users/rigofunc/Documents/sample.xlsx";

        // save to excel file
        reports.ToExcel(excelFile);

        // load from excel
        var loadFromExcel = Excel.Load<Report>(excelFile);
    }
}
```