# Functionality
The extensions of NPOI, which use attributes to control how to save IEnumerable&lt;T&gt; to excel.
- Use attribute to control excel column name, and cell index; 
- Use attribute to control excel cells SUM and cells MERGE behaviors; 
- Use attribute to control excel filter behaviors 
- Use attribute to control excel freeze behaviors

# How to use
1. Install NPOI.Extension by nuget

        PM> Install-Package NPOI.Extension
	
2. Reference NPOI.Extension in code

        using NPOI.Extension;
	
3. Apply attribute to the specified entity

        [Filter(FirstCol = 0, FirstRow = 0, LastCol = 2)]
        [Freeze(ColSplit = 2, RowSplit = 1, LeftMostColumn = 2, TopRow = 1)]
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
            [Column(Index = 6, Title = "佣金(元)", AllowSum = true)]
            public decimal Brokerage { get; set; }
            [Column(Index = 7, Title = "收益(元)", AllowSum = true)]
            public decimal Profits { get; set; }
        }

4. Using extension methods

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
        reports.ToExcel(@"C:\demo.xls");

5. NPOI.Extension demo

![NPOI.Extension demo](images/demo.PNG)
