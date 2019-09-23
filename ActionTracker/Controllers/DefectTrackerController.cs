using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Web.Http;

namespace ActionTracker.Controllers
{
    public class DefectTrackerController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);
        public IDbConnection connBI = new SqlConnection(ConfigurationManager.ConnectionStrings["BI"].ConnectionString);
        public IDbConnection connEDW = new SqlConnection(ConfigurationManager.ConnectionStrings["EDW"].ConnectionString);

        [HttpGet]
        public List<CT_QualityReport_Defect_DailyData> GetDailyData()
        {
            string sql = "select * from CT_QualityReport_Defect_DailyData";
            return conn.Query<CT_QualityReport_Defect_DailyData>(sql).ToList();
        }

        [HttpPost]
        public int UpdateDailyData(CT_QualityReport_Defect_DailyData data)
        {
            string sql = "update CT_QualityReport_Defect_DailyData set RootCause = @RootCause, RootCause1 = @RootCause1, CorrectiveActions = @CorrectiveActions, CorrectiveActions1 = @CorrectiveActions1, Owner1 = @Owner1, Owner2 = @Owner2, DueDate = @DueDate, Status = @Status, QAValidate = @QAValidate where ID = @ID";
            return conn.Execute(sql, data);
        }

        [HttpGet]
        public List<User> GetUserList()
        {
            string sql = "select * from EDW_CleanUser where BU = 'BU7' and Role in ('PE', 'EE')";
            return connEDW.Query<User>(sql).ToList();
        }

        [HttpPost]
        public List<SF_QualityReport_Defect_Daily> GetDefectDetail(CT_QualityReport_Defect_DailyData data)
        {
            string sql = "";
            if (data.Type.Trim() == "1")
            {
                sql = @"select * from [suznt060\sqlff6].FF_BI.[dbo].SF_QualityReport_Defect_Daily
                        where reportstation in (select  distinct(reportstation)  from [suznt060\sqlff6].FF_BI.[dbo].SF_QualityReport_Defect_Daily)
                        and project = 'TDN'and Component = @Component and FFStation = @Station and ProductionOrderNumber = @WO and TimeBegin = @TimeBegin and DefectDesc = @Defect";
            }
            if (data.Type.Trim() == "2")
            {
                sql = @"select * from [suznt060\sqlff6].FF_BI.[dbo].SF_QualityReport_Defect_Daily
                        where reportstation in (select  distinct(reportstation)  from [suznt060\sqlff6].FF_BI.[dbo].SF_QualityReport_Defect_Daily)
                        and project = 'TDN'and Component = @Component and FFStation = @Station and ProductionOrderNumber = @WO and TimeBegin = @TimeBegin";
            }
            if (data.Type.Trim() == "3")
            {
                sql = "select * from SF_FF_Repair_Result_O where TimeBegin = @TimeBegin and TestStation = @Station and Defect = @Defect and ProductionOrderNumber = @WO";
            }
            return connBI.Query<SF_QualityReport_Defect_Daily>(sql, data).ToList();
        }

        [HttpPost]
        public string SendMail(CT_QualityReport_Defect_DailyData data)
        {
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress("Action.Tracker@flex.com");
            mess.Subject = string.Format("Action Tracker");
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
            User user1 = GetUserList().Where(x => x.Name == data.Owner1).SingleOrDefault();
            User user2 = GetUserList().Where(x => x.Name == data.Owner2).SingleOrDefault();
            if (user1 != null)
            {
                mess.To.Add(user1.Mail);
            }
            if (user2 != null)
            {
                mess.To.Add(user2.Mail);
            }
            string str = "You have a new process issue, WO:" + data.WO + ", LLA PN: " + data.LLAPN + ", Component: " + data.Component + ", Defect: " + data.Defect + ". pls investigate immediately! You can visit: http://suznt004:8082/actiontracker";
            mess.Body = str.ToString();
            try
            {
                client.Send(mess);
                return "OK";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
