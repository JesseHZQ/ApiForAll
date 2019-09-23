using System;

namespace ApiForFCTKB.Controllers
{
    public class CheckListItem
    {
        public DateTime Date { get; set; }
        public string ListBy { get; set; }
        public string CheckItemDescription { get; set; }
        public string SystemId { get; set; }
        public string SystemName { get; set; }
    }
}