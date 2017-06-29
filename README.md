# IMPORTAMT
The future features will be support by [FluentExcel](https://github.com/Arch/FluentExcel), and will only support `Fluent API`.

The extensions for the NPOI, which provides IEnumerable&lt;T&gt; have save to and load from excel functionalities.

# Features
- [x] Decouple the configuration from the POCO model by using `fluent api`.
- [x] Support attributes based configuration.
- [x] Support POCO, so that if your mother langurage is Engilish, none any configurations;

The first two features will be very useful for English not their mother language developers.

# Overview

![NPOI.Extension demo](images/demo.PNG)

# Get Started

The following demo codes come from [sample](samples), download and run it for more information.

## Using Package Manager Console to install NPOI.Extension

        PM> Install-Package NPOI.Extension
    
## Reference NPOI.Extension in code

        using NPOI.Extension;
    
## Configure model's excel behaviors

We can use `fluent api` or `attributes` to configure the model excel behaviors. If both had been used, `fluent` configurations will has the `Hight Priority`

### 1. Use Fluent Api

```csharp
        public class Report {
            public string City { get; set; }
            public string Building { get; set; }
            public DateTime HandleTime { get; set; }
            public string Broker { get; set; }
            public string Customer { get; set; }
            public string Room { get; set; }
            public decimal Brokerage { get; set; }
            public decimal Profits { get; set; }
        }

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        static void FluentConfiguration() 
        {
            var fc = Excel.Setting.For<Report>();

            fc.HasStatistics("合计", "SUM", 6, 7)
              .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
              .HasFreeze(columnSplit: 2,rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

            fc.Property(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled();

            fc.Property(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled();

            fc.Property(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间")
              .HasDataFormatter("yyyy-MM-dd HH:mm:ss");
            
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
              .HasExcelTitle("佣金(元)");

            fc.Property(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
```

### 2. Use attributes

```csharp
    [Statistics(Name = "合计", Formula = "SUM", Columns = new[] { 6, 7 })]
    [Filter(FirstCol = 0, FirstRow = 0, LastCol = 2)]
    [Freeze(ColSplit = 2, RowSplit = 1, LeftMostColumn = 2, TopRow = 1)]
    public class Report {
        [Column(Index = 0, Title = "城市", AllowMerge = true)]
        public string City { get; set; }
        [Column(Index = 1, Title = "楼盘", AllowMerge = true)]
        public string Building { get; set; }
        [Column(Index = 2, Title = "成交时间", Formatter = "yyyy-MM-dd HH:mm:ss")]
        public DateTime HandleTime { get; set; }
        [Column(Index = 3, Title = "经纪人")]
        public string Broker { get; set; }
        [Column(Index = 4, Title = "客户")]
        public string Customer { get; set; }
        [Column(Index = 5, Title = "房源")]
        public string Room { get; set; }
        [Column(Index = 6, Title = "佣金(元)")]
        public decimal Brokerage { get; set; }
        [Column(Index = 7, Title = "收益(元)")]
        public decimal Profits { get; set; }
    }
```

## Export POCO to excel & Load IEnumerable&lt;T&gt; from excel.

```csharp
using System;
using NPOI.Extension;

namespace samples
{
    class Program
    {
        static void Main(string[] args)
        {
            // global call this
            FluentConfiguration();

            var len = 10;
            var reports = new Report[len];
            for (int i = 0; i < len; i++)
            {
                reports[i] = new Report
                {
                    City = "ningbo",
                    Building = "世茂首府",
                    HandleTime = new DateTime(2015, 11, 23),
                    Broker = "rigofunc 18957139**7",
                    Customer = "rigofunc 18957139**7",
                    Room = "2#1703",
                    Brokerage = 125M,
                    Profits = 25m
                };
            }

            var excelFile = @"/Users/rigofunc/Documents/sample.xlsx";

            // save to excel file
            reports.ToExcel(excelFile);

            // load from excel
            var loadFromExcel = Excel.Load<Report>(excelFile);
        }

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        static void FluentConfiguration() 
        {
            var fc = Excel.Setting.For<Report>();

            fc.HasStatistics("合计", "SUM", 6, 7)
              .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
              .HasFreeze(columnSplit: 2,rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

            fc.Property(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled();

            fc.Property(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled();

            fc.Property(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间")
              .HasDataFormatter("yyyy-MM-dd");
            
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
              .HasExcelTitle("佣金(元)");

            fc.Property(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
    }
}
 ```       
