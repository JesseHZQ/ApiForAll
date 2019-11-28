using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class SlotPlan
    {
        public string Slot { get; set; }
        public string SO { get; set; }
        public string MRP { get; set; }
        public string PD { get; set; }
        public string CORC { get; set; }
    }
}