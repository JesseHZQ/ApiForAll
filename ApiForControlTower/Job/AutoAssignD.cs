using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ApiForControlTower.Controllers;
using Quartz;
using Quartz.Impl;

namespace ApiForControlTower.Job
{
    public class AutoAssignD: IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            WorkAssignController wa = new WorkAssignController();
            wa.AutoAssignMorning();
        }
    }
}