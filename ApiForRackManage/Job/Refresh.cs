﻿using ApiForRackManage.Controllers;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForRackManage.Job
{
    public class Refresh : IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            SystemController system = new SystemController();
            system.RefreshQTY();
            system.JRefreshQTY();
        }
    }
}