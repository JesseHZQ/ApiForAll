using ApiForControlTower.Job;
using System;
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

namespace ApiForControlTower
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();

            // 执行计划任务
            JobScheduler.Start();

            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }

        protected void Application_End(object sender, EventArgs e)
        {
            //下面的代码是关键，可解决IIS应用程序池自动回收的问题
            Thread.Sleep(1000);
            //这里设置你的web地址，可以随便指向你的任意一个aspx页面甚至不存在的页面，目的是要激发Application_Start
            string url = "http://suznt004:8082/ControlTower/";
            HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
            Stream receiveStream = myHttpWebResponse.GetResponseStream();//得到回写的字节流
        }
    }
}
