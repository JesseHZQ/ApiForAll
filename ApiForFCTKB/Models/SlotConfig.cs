using System;

namespace ApiForFCTKB.Controllers
{
    public class SlotConfig
    {
        public int ID { get; set; }
        public string Slot { get; set; }
        public string PN { get; set; }
        public string Description { get; set; }
        public string QTY { get; set; }
        public string DelayTips { get; set; }
        public bool IsReady { get; set; }
        public DateTime LastUpdatedTime { get; set; }
    }
}
