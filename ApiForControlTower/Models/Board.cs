using System;

namespace ApiForControlTower.Controllers
{
    public class Board
    {
        public int ID { get; set; }
        public string SN { get; set; }
        public string Location { get; set; }
        public string NextLocation { get; set; }
        public string Station { get; set; }
        public DateTime Date { get; set; }
        public string Owner { get; set; }
        public string Remark { get; set; }
        public string Details { get; set; }

    }
}