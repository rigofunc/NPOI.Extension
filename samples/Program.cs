using System;
using System.IO;
using NPOI.Extension;

namespace samples {
    class Program {
        static void Main(string[] args) {
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

            // save to excel file
            reports.ToExcel(@"D:\demo.xls");

            var files = Directory.GetFiles(@"D:\excels", "*.xls", SearchOption.TopDirectoryOnly);
            foreach (var file in files) {
                // load from excel
                var loadFromExcel = Excel.Load<Model>(file);
            }
        }
    }
}
