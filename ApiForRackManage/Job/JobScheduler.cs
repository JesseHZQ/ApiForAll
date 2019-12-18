using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForRackManage.Job
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
            //创建触发器
            ITrigger triggerRefresh = TriggerBuilder.Create().WithIdentity("TimeTriggerRefresh", "TimeGroupRefresh").WithCronSchedule("0 0/5 * * * ?").Build();
            scheduler.ScheduleJob(jobRefresh, triggerRefresh);
            /*-------------计划任务代码实现------------------*/

            //启动
            scheduler.Start();
        }
    }
}