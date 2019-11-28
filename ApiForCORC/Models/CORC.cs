using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Models
{
    public class CORC
    {
        public string SO { get; set; }
        public CORCInfo CorcInfo { get; set; }
        public SOInfo SoInfo { get; set; }
    }
}