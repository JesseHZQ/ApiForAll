using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForFCTKB.Job
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
            IJobDetail jobRefresh = JobBuilder.Create<Refresh>().Build();
            IJobDetail jobMail = JobBuilder.Create<Mail>().Build();
            //创建触发器
            ITrigger triggerRefresh = TriggerBuilder.Create().WithIdentity("TimeTriggerRefresh", "TimeGroupRefresh").WithCronSchedule("0 0/30 8-17 * * ?").Build();
            ITrigger triggerMail = TriggerBuilder.Create().WithIdentity("TimeTriggerMail", "TimeGroupMail").WithCronSchedule("0 0 13 ? * MON-FRI").Build();
            scheduler.ScheduleJob(jobRefresh, triggerRefresh);
            scheduler.ScheduleJob(jobMail, triggerMail);
            /*-------------计划任务代码实现------------------*/

            //启动
            scheduler.Start();
        }
    }
}