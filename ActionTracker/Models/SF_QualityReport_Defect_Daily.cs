    

using System;

namespace ActionTracker
{
    public partial class SF_QualityReport_Defect_Daily
    {
       public int ID { get; set; }
       public string BU { get; set; }
       public string Project { get; set; }
       public string Level { get; set; }
       public string ReportStation { get; set; }
       public string FFStationType { get; set; }
       public DateTime? TimeBegin { get; set; }
       public DateTime? TimeEnd { get; set; }
       public string SN { get; set; }
       public string ProductionOrderNumber { get; set; }
       public string FFStation { get; set; }
       public string Family { get; set; }
       public string PartNumber { get; set; }
       public string PartDesc { get; set; }
       public string Failure { get; set; }
       public string DefectCode { get; set; }
       public string DefectDesc { get; set; }
       public string Location { get; set; }
       public string Component { get; set; }
       public string Remark { get; set; }
       public string UserID { get; set; }
       public DateTime? Time { get; set; }
       public DateTime InsertTime { get; set; }
    }

}