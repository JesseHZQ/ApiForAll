namespace ApiForControlTower.Controllers
{
    public class PCBAShortage
    {
        public int ID { get; set; }
        public string Slot { get; set; }
        public string PN { get; set; }
        public string Description { get; set; }
        public string SupplierName { get; set; }
        public int QTY { get; set; }
        public string Buyer { get; set; }
        public string ETA { get; set; }
        public int IsReceived { get; set; }
        public string MRP { get; set; }
    }
}