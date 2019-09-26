using System.Collections.Generic;

namespace ActionTracker.Controllers
{
    public class Resp
    {
        public List<CT_QualityReport_Defect_DailyData> list { get; set; }
        public int total { get; set; }
    }
}