using System;

namespace ApiForControlTower.Controllers
{
    public class TempHumid
    {
        public DateTime CreationTime { get; set; }
        public double Temperature { get; set; }
        public double Humidity { get; set; }
        public string Location { get; set; }
    }
}