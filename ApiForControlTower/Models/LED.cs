using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class LED
    {
        public int Id { get; set; }
        public int Item { get; set; }
        public string BayName { get; set; }
        public string SystemName { get; set; }
        public string Station { get; set; }
        public string ShippingTime { get; set; }
        public string IP { get; set; }
        public string PlanShipDate { get; set; }
        public string ShippingDate { get; set; }
        public string Slot { get; set; }
        public string Type { get; set; }
        public string Model { get; set; }
        public string Customer { get; set; }
        public string PO { get; set; }
        public string SO { get; set; }
        public string MRP { get; set; }
        public string PD { get; set; }
        public string CORC { get; set; }
        public string Launch { get; set; }
        public string CoreBU { get; set; }
        public string PV { get; set; }
        public string OI { get; set; }
        public string TestBU { get; set; }
        public string CSW { get; set; }
        public string QFAA { get; set; }
        public string BU { get; set; }
        public string Pack { get; set; }
        public string CORC_Issue { get; set; }
        public string Engineering_Issue { get; set; }
        public string Remark { get; set; }
        public bool IsLoad { get; set; }
        public string LoadInfo { get; set; }
        public string GroupNum { get; set; }
    }
}