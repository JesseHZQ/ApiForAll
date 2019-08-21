namespace ApiForControlTower.Models
{
    public class SlotPlan
    {
        public string Slot { get; set; }
        public string Type { get; set; }
        public string Model { get; set; }
        public string Customer { get; set; }
        public string PO { get; set; }
        public string SO { get; set; }
        public string MRP { get; set; }
        public string PD { get; set; }
        public string CORC { get; set; }
        public string CORC_Issue { get; set; }
        public string Engineering_Issue { get; set; }
        public string Remark { get; set; }
    }
}