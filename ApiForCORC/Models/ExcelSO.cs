using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class ExcelSO
    {
        public string PDWeek { get; set; }
        public string Slot { get; set; }
        public string SO { get; set; }
        public string Customer { get; set; }
        public string ShipToLocation { get; set; }
        public string Incoterms { get; set; }
        public string Week { get; set; }
        public string Issues { get; set; }
        public string DueDateToBeFixed { get; set; }
        public string RemainFixedDay { get; set; }
        public string OverDueWorkingDay { get; set; }
        public string EarliestShipDay { get; set; }
        public string Impacts { get; set; }
        public DateTime InsertTime { get; set; }
        public bool IsShow { get; set; }
    }
}