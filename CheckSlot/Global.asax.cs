﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace CheckSlot
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            //System.Timers.Timer myTimer = new System.Timers.Timer(1000 * 30);
            //myTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
            //myTimer.Enabled = true;
            //myTimer.AutoReset = true;
        }

        //void Session_End(object sender, EventArgs e)
        //{
        //    //下面的代码是关键，可解决IIS应用程序池自动回收的问题
        //    System.Threading.Thread.Sleep(1000);
        //    //触发事件, 写入提示信息
        //    MailTask.SendMail();
        //    //这里设置你的web地址，可以随便指向你的任意一个aspx页面甚至不存在的页面，目的是要激发Application_Start
        //    //使用您自己的URL
        //    string url = "http://suznt004:3001/automail";
        //    System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
        //    System.Net.HttpWebResponse myHttpWebResponse = (System.Net.HttpWebResponse)myHttpWebRequest.GetResponse();
        //    System.IO.Stream receiveStream = myHttpWebResponse.GetResponseStream();//得到回写的字节流

        //    // 在会话结束时运行的代码。
        //    // 注意: 只有在 Web.config 文件中的 sessionstate 模式设置为 InProc 时，才会引发 Session_End 事件。
        //    // 如果会话模式设置为 StateServer
        //    // 或 SQLServer，则不会引发该事件。
        //}

        ////protected void Application_End(object sender, EventArgs e)
        ////{
        ////    //下面的代码是关键，可解决IIS应用程序池自动回收的问题 
        ////    Thread.Sleep(1000);
        ////    //这里设置你的web地址，可以随便指向你的任意一个aspx页面甚至不存在的页面，目的是要激发Application_Start 
        ////    //我这里是一个验证token的地址。
        ////    string url = @"http://suznt004:8086/syskb/";
        ////    //string url = System.Configuration.ConfigurationManager.AppSettings["tokenurl"];
        ////    HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
        ////    HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
        ////    Stream receiveStream = myHttpWebResponse.GetResponseStream();//得到回写的字节流 
        ////}

        //private static void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        //{
        //    //if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
        //    //{
        //    //    if (DateTime.Now.Hour == 16 && DateTime.Now.Minute == 30)
        //    //    {
        //    //        MailTask.SendMail();
        //    //    }
        //    //}
        //    //if (DateTime.Now.DayOfWeek == DayOfWeek.Thursday)
        //    //{
        //    //    if (DateTime.Now.Hour == 08)
        //    //    {
        //    MailTask.SendMail();
        //    //    }
        //    //}

        //    //MailTask.SendMail();

        //}
    }
}
