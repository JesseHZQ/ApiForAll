using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ApiForFCTKBNew.Controllers;
using Quartz;
using Quartz.Impl;

namespace ApiForFCTKBNew.Job
{
    public class Refresh : IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            RefreshController refresh = new RefreshController();
            refresh.Refresh();

            SupermarketController supermarket = new SupermarketController();
            supermarket.RefreshQTY();
        }
    }
}