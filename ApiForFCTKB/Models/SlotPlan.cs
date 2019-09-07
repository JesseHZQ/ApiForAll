﻿using System;
using System.Collections.Generic;

namespace ApiForFCTKB.Controllers
{
    public class SlotPlan
    {
        public int ID { get; set; }
        public string PlanShipDate { get; set; }
        public string ShippingDate { get; set; }
        public string ShippingType { get; set; }
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
        public DateTime LastUpdatedTime { get; set; }
        public List<SlotConfig> ConfigList { get; set; }
        public List<SlotShortage> ShortageList { get; set; }
        public List<SlotEIssue> EIssueList { get; set; }
    }
}
