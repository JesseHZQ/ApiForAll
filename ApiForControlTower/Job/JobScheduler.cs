﻿using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Job
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
            IJobDetail jobD = JobBuilder.Create<AutoAssignD>().Build();
            IJobDetail jobS = JobBuilder.Create<AutoAssignS>().Build();
            IJobDetail jobCORCRemind = JobBuilder.Create<CorcRemind>().Build();
            //创建触发器
            ITrigger triggerD = TriggerBuilder.Create().WithIdentity("TimeTriggerD", "TimeGroupD").WithCronSchedule("0 38 8 ? * MON-FRI").Build();
            ITrigger triggerS = TriggerBuilder.Create().WithIdentity("TimeTriggerS", "TimeGroupS").WithCronSchedule("0 35 15 ? * MON-FRI").Build();
            ITrigger triggerCORCRemind = TriggerBuilder.Create().WithIdentity("TimeTriggerCORCRemind", "TimeGroupCORCRemind").WithCronSchedule("0 55 14 ? * MON-FRI").Build();
            scheduler.ScheduleJob(jobD, triggerD);
            scheduler.ScheduleJob(jobS, triggerS);
            scheduler.ScheduleJob(jobCORCRemind, triggerCORCRemind);
            /*-------------计划任务代码实现------------------*/

            //启动
            scheduler.Start();
        }
    }
}