using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class Work
    {
        public int Id { get; set; }
        public string BayName { get; set; }
        public string SystemName { get; set; }
        public string Station { get; set; }
        public DateTime Date { get; set; }
        public int OwnerId { get; set; }
    }
}