using System;

namespace ApiForFCTKB.Controllers
{
    public partial class SystemController
    {
        public class SlotCORC
        {
            public string Slot { get; set; }
            public string CORC { get; set; }
            public string CORC_Issue { get; set; }
            public DateTime InsertTime { get; set; }
            public DateTime LastUpdateTime { get; set; }
            public DateTime CloseTime { get; set; }
        }
    }
}
