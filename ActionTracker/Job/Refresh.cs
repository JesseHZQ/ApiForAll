using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Web;
using ActionTracker.Controllers;
using Dapper;
using Quartz;
using Quartz.Impl;

namespace ActionTracker.Job
{
    public class Refresh: IJob
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);
        public IDbConnection connEDW = new SqlConnection(ConfigurationManager.ConnectionStrings["EDW"].ConnectionString);
        public void Execute(IJobExecutionContext context)
        {
            string sql = "select * from CT_QualityReport_Defect_DailyData";
            List<CT_QualityReport_Defect_DailyData> list = conn.Query<CT_QualityReport_Defect_DailyData>(sql).ToList();

            string sqlUser = "select * from EDW_CleanUser where BU = 'BU7' and Role in ('PE', 'EE')";
            List<User> listUser = connEDW.Query<User>(sqlUser).ToList();

            list = list.Where(x => new TimeSpan(DateTime.Now.Ticks).Subtract(new TimeSpan(DateTime.Parse(x.DueDate).Ticks)).Duration().Days > 3
            && (x.Owner1 != "" && x.Owner1 != null) && (x.RootCause == "" || x.RootCause == null)).ToList();

            foreach (CT_QualityReport_Defect_DailyData item in list)
            {
                User user = listUser.Where(x => x.Name == item.Owner1).SingleOrDefault();
                switch (user.Name)
                {
                    case "Xuefei Ji":
                    case "Liming Zhu":
                    case "QF Xu":
                        MailMessage mess = new MailMessage();
                        mess.From = new MailAddress("Action.Tracker@flex.com");
                        mess.Subject = string.Format("Action Tracker");
                        mess.IsBodyHtml = true;
                        mess.BodyEncoding = System.Text.Encoding.UTF8;
                        SmtpClient client = new SmtpClient();
                        client.Host = "10.194.51.14";
                        client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
                        mess.To.Add("zhenghua.zhou@flex.com");
                        string str = "Action Owner: " + user.Name + " don't response for 3 days, You can visit: http://suznt004:8082/actiontracker to view more details";
                        mess.Body = str.ToString();
                        client.Send(mess);
                        break;
                    case "Hunter Ge":
                    case "Nick Qin":
                        MailMessage mess1 = new MailMessage();
                        mess1.From = new MailAddress("Action.Tracker@flex.com");
                        mess1.Subject = string.Format("Action Tracker");
                        mess1.IsBodyHtml = true;
                        mess1.BodyEncoding = System.Text.Encoding.UTF8;
                        SmtpClient client1 = new SmtpClient();
                        client1.Host = "10.194.51.14";
                        client1.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
                        mess1.To.Add("jerry.zhang@flex.com");
                        string str1 = "Action Owner: " + user.Name + " don't response for 3 days, You can visit: http://suznt004:8082/actiontracker to view more details";
                        mess1.Body = str1.ToString();
                        client1.Send(mess1);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}