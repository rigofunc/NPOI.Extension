namespace samples {
    using NPOI.Extension;

    public class Model {
        [Column(Index = 0)]
        public int ApplyID { get; set; }
        [Column(Index = 1)]
        public string CityName { get; set; }
        [Column(Index = 2)]
        public string BuildingName { get; set; }
        [Column(Index = 3)]
        public string CustomerName { get; set; }
        [Column(Index = 4)]
        public string CustomerPhone { get; set; }
        [Column(Index = 5)]
        public string Room { get; set; }
        [Column(Index = 6)]
        public string HandleTime { get; set; }
        [Column(Index = 7)]
        public string Price { get; set; }
        [Column(Index = 8)]
        public string BrokerName { get; set; }
        [Column(Index = 9)]
        public string BrokerPhone { get; set; }
        [Column(Index = 10)]
        public string Brokerage { get; set; }
        [Column(Index = 11)]
        public string Profits { get; set; }
        [Column(Index = 1)]
        public string SaaSMoney { get; set; }
        [Column(Index = 12)]
        public string Subsidies { get; set; }
        [Column(Index = 13)]
        public string BrokerInvoiceTime { get; set; }
        [Column(Index = 14)]
        public string KaKaoInvoiceTime { get; set; }
        [Column(Index = 15)]
        public string InvoiceCode { get; set; }
        [Column(Index = 16)]
        public string InvoiceMoney { get; set; }
    }
}
