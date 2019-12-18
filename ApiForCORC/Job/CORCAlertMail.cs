using ApiForCORC.Controllers;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Job
{
    public class CORCAlertMail: IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            ExcelCORCController ec = new ExcelCORCController();
            ec.GetExcelCORC();

            ExcelSOController es = new ExcelSOController();
            es.GetExcelSO();
        }
    }
}