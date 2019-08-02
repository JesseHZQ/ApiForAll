using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class Failure
    {
        public int Id { get; set; }
        public string FailureStation { get; set; }
        public string FailureType { get; set; }
        public string FailureDesc { get; set; }
    }
}