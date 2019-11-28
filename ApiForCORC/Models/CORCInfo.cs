using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class CORCInfo
    {
        public string SO { get; set; }
        public string Status { get; set; }
        public string Customer { get; set; }
        public string SO_Verify { get; set; }
        public string ProductComments { get; set; }
        public string FactoryRequirement { get; set; }
        public int SiteVoltage { get; set; }
        public string DUTOrientation { get; set; }
        public string ComputerEquipped { get; set; }
        public string IGXLVersion { get; set; }
        public string IPAddress { get; set; }
        public string OtherSpecial { get; set; }
        public string SpecialComments { get; set; }
    }
}