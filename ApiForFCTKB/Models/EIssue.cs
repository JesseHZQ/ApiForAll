using System;

namespace ApiForFCTKB.Controllers
{
    public class SlotEIssue
    {
        public int Id { get; set; }
        public string Slot { get; set; }
        public string Item { get; set; }
        public DateTime Date { get; set; }
        public string Status { get; set; }
        public DateTime LastUpdatedDate { get; set; }
    }
}