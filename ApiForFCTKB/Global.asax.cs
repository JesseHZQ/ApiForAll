﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace ApiForFCTKB
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

            System.Timers.Timer myTimer = new System.Timers.Timer(1000 * 60);
            myTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
            myTimer.Enabled = true;
            myTimer.AutoReset = true;
        }


        protected void Session_End(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(1000);
            string url = @"http://www.baidu.com";
            System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
            System.Net.HttpWebResponse myHttpWebResponse = (System.Net.HttpWebResponse)myHttpWebRequest.GetResponse();
            System.IO.Stream receiveStream = myHttpWebResponse.GetResponseStream();
        }

        private static void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        {
            if (DateTime.Now.DayOfWeek != DayOfWeek.Saturday && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
            {
                if (DateTime.Now.Hour == 13 && DateTime.Now.Minute == 0)
                {
                    MailTask mt = new MailTask();
                    mt.SendMail();
                }
            }
        }
    }
}
