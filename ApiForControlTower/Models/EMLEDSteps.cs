using System;

namespace ApiForControlTower.Controllers
{
    public class EMLEDSteps
    {
        public int ID { get; set; }
        public int Aslog { get; set; }
        public string AssyStep { get; set; }
        public DateTime EndTime { get; set; }
    }
}