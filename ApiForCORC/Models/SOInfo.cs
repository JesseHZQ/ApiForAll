using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class SOInfo
    {
        public string SO { get; set; }
        public List<SOLine> SOLineList { get; set; }
    }
}