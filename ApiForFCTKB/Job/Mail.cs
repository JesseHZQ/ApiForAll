using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ApiForFCTKB;
using ApiForFCTKB.Controllers;
using Quartz;
using Quartz.Impl;

namespace ApiForFCTKB.Job
{
    public class Mail : IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            MailTask mt = new MailTask();
            mt.SendMail();
        }
    }
}