using Aspose.Cells;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Web;
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
        public Resp GetDailyData(CT_QualityReport_Defect_DailyData data)
        {
            int start = (data.Index - 1) * data.Size;
            int length = data.Size;
            string sql, sqlTotal;
            if (data.ActionOwner == "" || data.ActionOwner == null)
            {
                sql = "select * from CT_QualityReport_Defect_DailyData where QEOwner like ('%" + data.QEOwner + "%') and WO like ('%" + data.WO + "%') and LLAPN like ('%" + data.LLAPN + "%') and Component like ('%" + data.Component + "%') order by TIMEBEGIN DESC OFFSET " + start + " ROWS FETCH NEXT " + length + " ROWS ONLY";
                sqlTotal = "select count(*) as total from CT_QualityReport_Defect_DailyData where QEOwner like ('%" + data.QEOwner + "%') and WO like ('%" + data.WO + "%') and LLAPN like ('%" + data.LLAPN + "%') and Component like ('%" + data.Component + "%')";

            }
            else
            {
                sql = "select * from CT_QualityReport_Defect_DailyData where (Owner1 like ('%" + data.ActionOwner + "%') or Owner2 like ('%" + data.ActionOwner + "%')) and QEOwner like ('%" + data.QEOwner + "%') and WO like ('%" + data.WO + "%') and LLAPN like ('%" + data.LLAPN + "%') and Component like ('%" + data.Component + "%') order by TIMEBEGIN DESC OFFSET " + start + " ROWS FETCH NEXT " + length + " ROWS ONLY";
                sqlTotal = "select count(*) as total from CT_QualityReport_Defect_DailyData where (Owner1 like ('%" + data.ActionOwner + "%') or Owner2 like ('%" + data.ActionOwner + "%')) and QEOwner like ('%" + data.QEOwner + "%') and WO like ('%" + data.WO + "%') and LLAPN like ('%" + data.LLAPN + "%') and Component like ('%" + data.Component + "%')";
            }
            Resp resp = new Resp();
            resp.list = conn.Query<CT_QualityReport_Defect_DailyData>(sql).ToList();
            resp.total = conn.Query<int>(sqlTotal).SingleOrDefault();
            return resp;
        }

        [HttpPost]
        public int UpdateDailyData(CT_QualityReport_Defect_DailyData data)
        {
            string sql = "update CT_QualityReport_Defect_DailyData set RootCause = @RootCause, RootCause1 = @RootCause1, CorrectiveActions = @CorrectiveActions, CorrectiveActions1 = @CorrectiveActions1, Owner1 = @Owner1, Owner2 = @Owner2, DueDate = @DueDate, Status = @Status, QAValidate = @QAValidate where ID = @ID";
            return conn.Execute(sql, data);
        }

        [HttpGet]
        public List<string> GetWOList(string WO)
        {
            string sql = "select distinct WO from CT_QualityReport_Defect_DailyData where WO like '%" + WO + "%'";
            return conn.Query<string>(sql).ToList();
        }
        [HttpGet]
        public List<string> GetLLAPNList(string LLAPN)
        {
            string sql = "select distinct LLAPN from CT_QualityReport_Defect_DailyData where LLAPN like '%" + LLAPN + "%'";
            return conn.Query<string>(sql).ToList();
        }
        [HttpGet]
        public List<User> GetUserNameList(string Name)
        {
            string sql = "select * from EDW_CleanUser where BU = 'BU7' and Role in ('PE', 'EE', 'Production-PTH') and Name like '%" + Name + "%'";
            return connEDW.Query<User>(sql).ToList();
        }
        [HttpGet]
        public List<string> GetComponentList(string Component)
        {
            string sql = "select distinct Component from CT_QualityReport_Defect_DailyData where Component like '%" + Component + "%'";
            return conn.Query<string>(sql).ToList();
        }
        [HttpGet]
        public void generateExcel()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\ActionTracker\file\Action Tracker.xlsx");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 0;
            List<CT_QualityReport_Defect_DailyData> list = GetDailyData();
            foreach (CT_QualityReport_Defect_DailyData item in list)
            {
                y++;
                cells[y, 0].PutValue(item.Date);
                cells[y, 1].PutValue(item.RepeatIssue);
                cells[y, 2].PutValue(item.WO);
                cells[y, 3].PutValue(item.LLAPN);
                cells[y, 4].PutValue(item.Component);
                cells[y, 5].PutValue(item.Defect);
                cells[y, 6].PutValue(item.Station);
                cells[y, 7].PutValue(item.DefectQty);
                cells[y, 8].PutValue(item.Yieldloss);
                cells[y, 9].PutValue(item.QEOwner);
                cells[y, 10].PutValue(item.RootCause);
                cells[y, 11].PutValue(item.CorrectiveActions);
                cells[y, 12].PutValue(item.Owner1);
                cells[y, 13].PutValue(item.RootCause1);
                cells[y, 14].PutValue(item.CorrectiveActions1);
                cells[y, 15].PutValue(item.Owner2);
                cells[y, 16].PutValue(item.DueDate);
                cells[y, 17].PutValue(item.Status);
                cells[y, 18].PutValue(item.QAValidate);
            }
            wb.Save(HttpContext.Current.Response, "ActionTracker_" + DateTime.Now.ToShortDateString() + ".xlsx", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Xlsx));
        }

        [HttpPost]
        public List<CT_QualityReport_Defect_DailyData> GetDailyDataBy(CT_QualityReport_Defect_DailyData data)
        {
            string sql = "select * from CT_QualityReport_Defect_DailyData where Defect like ('%" + data.Defect + "%') and LLAPN like ('%" + data.LLAPN + "%') and Component like ('%" + data.Component + "%')";
            return conn.Query<CT_QualityReport_Defect_DailyData>(sql, data).ToList();
        }

        [HttpGet]
        public List<User> GetUserList()
        {
            string sql = "select * from EDW_CleanUser where BU = 'BU7' and Role in ('PE', 'EE', 'Production-PTH')";
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
                        and project = 'TDN' and Component = @Component and TimeBegin = @TimeBegin and DefectDesc = @Defect";
            }
            if (data.Type.Trim() == "2")
            {
                sql = @"select * from [suznt060\sqlff6].FF_BI.[dbo].SF_QualityReport_Defect_Daily
                        where project = 'TDN' and Component = @Component and TimeBegin = @TimeBegin";
            }
            if (data.Type.Trim() == "3")
            {
                //sql = "select * from SF_FF_Repair_Result_O where TimeBegin = @TimeBegin and TestStation = @Station and Defect = @Defect and ProductionOrderNumber = @WO";
                //sql = "select * from SF_FF_Repair_Result_O where Id = @ReportID";
                sql = @"select * from [suznt060\sqlff6].FF_BI.dbo.SF_FF_Repair_Result_O   so
                        where  so.TimeBegin between (select convert(datetime,convert(varchar(10),'2019-11-12',120))) and (select convert(datetime,convert(varchar(10),getdate()-1,120)))
                        and so.PartNumber = @LLAPN and so.ComponentPart = @Component and so.Defect = @Defect and so.teststation = @Station";
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