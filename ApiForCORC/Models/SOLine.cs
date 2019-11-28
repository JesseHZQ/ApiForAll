using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class SOLine
    {
        public string Line { get; set; }
        public string OrderedItem { get; set; }
        public string Qty { get; set; }
        public string QtyShipped { get; set; }
        public string ItemDescription { get; set; }
        public string ScheduleShipDate { get; set; }
    }
}