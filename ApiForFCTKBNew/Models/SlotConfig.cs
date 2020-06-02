using System;

namespace ApiForFCTKBNew.Models
{
    public class SlotConfig
    {
        public int ID { get; set; }
        public string Slot { get; set; }
        public string PN { get; set; }
        public string Description { get; set; }
        public int QTY { get; set; }
        public string DelayTips { get; set; }
        public int IsReady { get; set; }
        public DateTime LastUpdatedTime { get; set; }
    }
}