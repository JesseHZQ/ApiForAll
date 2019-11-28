using ApiForControlTower.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class LED: SlotPlan
    {
        public int Id { get; set; }
        public int Item { get; set; }
        public string BayName { get; set; }
        public string SystemName { get; set; }
        public string Station { get; set; }
        public string ShippingTime { get; set; }
        public string IP { get; set; }
        public string VNCIP { get; set; }
        public DateTime LastUpdatedTime { get; set; }
        public int OwnerId { get; set; }
        public string UserName { get; set; }
        public string WeChatName { get; set; }
        public List<EIssue> EIssueList { get; set; }
    }
}