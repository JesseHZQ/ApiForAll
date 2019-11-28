using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForCORC.Job
{
    public class JobScheduler
    {
        public static void Start()
        {
            //调度器工厂
            ISchedulerFactory factory = new StdSchedulerFactory();

            //调度器
            IScheduler scheduler = factory.GetScheduler();
            scheduler.GetJobGroupNames();

            /*-------------计划任务代码实现------------------*/
            //创建任务
            IJobDetail job = JobBuilder.Create<CORCAlertMail>().Build();

            //创建触发器
            ITrigger trigger = TriggerBuilder.Create().WithIdentity("TriggerMail", "GroupMail").WithCronSchedule("0 30 16 ? * MON-FRI").Build();
            scheduler.ScheduleJob(job, trigger);
            /*-------------计划任务代码实现------------------*/

            //启动
            scheduler.Start();
        }
    }
}