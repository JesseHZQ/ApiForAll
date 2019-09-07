using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ApiForFCTKB.Controllers;
using Quartz;
using Quartz.Impl;

namespace ApiForFCTKB.Job
{
    public class Refresh: IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            SystemController system = new SystemController();
            system.Refresh();
        }
    }
}