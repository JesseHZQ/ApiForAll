using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class ExcelCORC
    {
        public string Slot { get; set; }
        public string SO { get; set; }
        public string Week { get; set; }
        public string MRPWeek { get; set; }
        public string CORCStatus { get; set; }
        public string CORCDueDateToBeFixed { get; set; }
        public string RemainFixedDay { get; set; }
        public string OverDueWorkingDay { get; set; }
        public string EarliestShipDay { get; set; }
        public string Impacts { get; set; }
        public DateTime InsertTime { get; set; }
        public bool IsShow { get; set; }
    }
}