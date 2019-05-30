using System.Collections.Generic;

namespace ApiForFCTKB.Controllers
{
    public partial class SystemController
    {
        public class Resp
        {
            public int code { get; set; }
            public List<SlotPlan> slotPlans { get; set; }
            public List<SlotConfig> slotConfigs { get; set; }
            public List<SlotShortage> slotShortages { get; set; }
            public string message { get; set; }
        }
    }
}
