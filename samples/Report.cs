using System;
using Arch.FluentExcel;

namespace samples
{
    [Statistics(Name = "合计", Formula = "SUM", Columns = new[] { 6, 7 })]
    [Filter(FirstCol = 0, FirstRow = 0, LastCol = 2)]
    [Freeze(ColSplit = 2, RowSplit = 1, LeftMostColumn = 2, TopRow = 1)]
    public class Report
    {
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
}
