using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class Task
    {
        public int Id { get; set; }
        public string BayName { get; set; }
        public string SystemName { get; set; }
        public string FailureStation { get; set; }
        public string FailureType { get; set; }
        public string FailureDesc { get; set; }
        public int OwnerId { get; set; }
        public string Status { get; set; }
        public string RootCause { get; set; }
        public string Solution { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime CloseTime { get; set; }
    }
}