using System;

namespace ApiForControlTower.Controllers
{
    public class CORC
    {
        public string Slot { get; set; }
        public string PD { get; set; }
        public string CORC_Issue { get; set; }
        public DateTime InsertTime { get; set; }
        public DateTime LastUpdateTime { get; set; }
        public DateTime CloseTime { get; set; }
    }
}