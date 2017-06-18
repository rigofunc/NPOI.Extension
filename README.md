The extensions for the NPOI, which provides IEnumerable&lt;T&gt; have save to and load from excel functionalities.

# Features
- [x] Decouple the configuration from the POCO model by using `fluent api`.
- [x] Support attributes based configuration.
- [x] Support POCO, so that if your mother langurage is Engilish, none any configurations;

The first two features will be very useful for English not their mother language developers.

# Overview

![NPOI.Extension demo](images/demo.PNG)

# Get Started
## Using Package Manager Console to install NPOI.Extension

        PM> Install-Package NPOI.Extension
    
## Reference NPOI.Extension in code

        using NPOI.Extension;
    
## Configure model's excel behaviors

We can use `fluent api` or `attributes` to configure the model excel behaviors.

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
        static void ExcelFluentConfig() 
        {
            var mc = Excel.Setting.For<Report>();

            mc.HasStatistics("合计", "SUM", 6, 7)
              .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
              .HasFreeze(columnSplit: 2,rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

            mc.Property(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled(true);

            mc.Property(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled(true);

            mc.Property(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间");
            
            mc.Property(r => r.Broker)
              .HasExcelIndex(3)
              .HasExcelTitle("经纪人");
            
            mc.Property(r => r.Customer)
              .HasExcelIndex(4)
              .HasExcelTitle("客户");

            mc.Property(r => r.Room)
              .HasExcelIndex(5)
              .HasExcelTitle("房源");

            mc.Property(r => r.Brokerage)
              .HasExcelIndex(6)
              .HasExcelTitle("佣金(元)");

            mc.Property(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
```

### 2. Use attributes

```csharp
        [Filter(FirstCol = 0, FirstRow = 0, LastCol = 2)]
        [Freeze(ColSplit = 2, RowSplit = 1, LeftMostColumn = 2, TopRow = 1)]
        [Statistics(Name = "合计", Formula = "SUM", Columns = new[] { 6, 7 })]
        public class Report {
            [Column(Index = 0, Title = "城市", AllowMerge = true)]
            public string City { get; set; }
            [Column(Index = 1, Title = "楼盘", AllowMerge = true)]
            public string Building { get; set; }
            [Column(Index = 2, Title = "成交时间")]
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

## Export POCO to excel.

```csharp
        var len = 1000;
        var reports = new Report[len];
        for (int i = 0; i < len; i++) {
            reports[i] = new Report {
                    City = "ningbo",
                    Building = "世茂首府",
                    HandleTime = new DateTime(2015, 11, 23),
                    Broker = "RigoFunc 18957139**7",
                    Customer = "RigoFunc 18957139**7",
                    Room = "2#1703",
                    Brokerage = 125M,
                    Profits = 25m
            };

            // other data here...
        }

        // save the excel file
        reports.ToExcel(@"C:\demo.xlsx");
 ```       
## Load IEnumerable&lt;T&gt; from excel

```csharp
        // load from excel
        var loadFromExcel = Excel.Load<Report>(@"C:\demo.xlsx");
```

## Custom excel export setting

The POCO export use following setting, so, the end user can costomize the setting like `Excel.Setting.DateFormatter = "yyyy-MM-dd";`

```csharp
    public class ExcelSetting
    {
        public string Company { get; set; } = "rigofunc (xuyingting)";

        public string Author { get; set; } = "rigofunc (xuyingting)";

        public string Subject { get; set; } = "The extensions of NPOI, which provides IEnumerable<T>; save to and load from excel.";

        public bool UserXlsx { get; set; } = true;

        public string DateFormatter { get; set; } = "yyyy-MM-dd HH:mm:ss";
    }
```