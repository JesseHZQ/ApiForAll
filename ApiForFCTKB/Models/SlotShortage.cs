using System;

namespace ApiForFCTKB.Controllers
{
    public class SlotShortage
    {
        public int ID { get; set; }
        public string Slot { get; set; }
        public string PN { get; set; }
        public string Description { get; set; }
        public string SupplierName { get; set; }
        public string QTY { get; set; }
        public string Buyer { get; set; }
        public string ETA { get; set; }
        public bool IsReceived { get; set; }
        public DateTime InsertTime { get; set; }
        public DateTime LastUpdatedTime { get; set; }
        public DateTime ReceivedTime { get; set; }
    }
}
