using System;

namespace ApiForControlTower.Controllers
{
    public class EIssue
    {
        public int Id { get; set; }
        public string Slot { get; set; }
        public string Item { get; set; }
        public DateTime Date { get; set; }
        public string Status { get; set; }
    }
}