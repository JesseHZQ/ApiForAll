using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForRackManage.Models
{
    public class ConfigBoard
    {
        public string ID { get; set; }
        public string Mark { get; set; }
        public string Slot { get; set; }
        public string Type { get; set; }
        public string PN { get; set; }
        public string SN { get; set; }
        public string SN2 { get; set; }
        public string RevD { get; set; }
        public string RevD2 { get; set; }
        public string RevL { get; set; }
        public string RevL2 { get; set; }
        public string CurrentTime { get; set; }
        public string Item { get; set; }
    }
}