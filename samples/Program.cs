using System;
using System.IO;
using NPOI.Extension;

namespace samples
{
    class Program
    {
        static void Main(string[] args)
        {
            var len = 1000;
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

                // other data here...
            }

            // save to excel file
            reports.ToExcel(@"D:\demo.xlsx");

            var files = Directory.GetFiles(@"D:\excels", "*.xlsx", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                // load from excel
                var loadFromExcel = Excel.Load<Model>(file);
            }
        }

		/// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
		/// </summary>
		static void FluentConfiguration() 
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
    }
}
