using ActionTracker.Controllers;
using System;
using System.Collections.Generic;

namespace ActionTracker
{
    public partial class CT_QualityReport_Defect_DailyData: PageClass
    {
        public int ID { get; set; }
        public string Date { get; set; }
        public string RepeatIssue { get; set; }
        public string WO { get; set; }
        public string LLAPN { get; set; }
        public string Component { get; set; }
        public string Defect { get; set; }
        public string Station { get; set; }
        public string DefectQty { get; set; }
        public string DefectRate { get; set; }
        public string Yieldloss { get; set; }
        public string RootCause { get; set; }
        public string CorrectiveActions { get; set; }
        public string RootCause1 { get; set; }
        public string CorrectiveActions1 { get; set; }
        public string Owner1 { get; set; }
        public string Owner2 { get; set; }
        public string QEOwner { get; set; }
        public string DueDate { get; set; }
        public string Status { get; set; }
        public string QAValidate { get; set; }
        public DateTime? INSERTTIME { get; set; }
        public DateTime TimeBegin { get; set; }
        public string Type { get; set; }
    }

}