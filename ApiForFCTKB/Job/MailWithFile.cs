using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Web;
using ApiForFCTKB;
using ApiForFCTKB.Controllers;
using ApiForFCTKB.Models;
using Dapper;
using Quartz;
using Quartz.Impl;

namespace ApiForFCTKB.Job
{
    public class MailWithFile : IJob
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public void Execute(IJobExecutionContext context)
        {
            //具体的方法
            SystemController sc = new SystemController();
            string UFPath = sc.generateExcel("UF");
            string IFPath = sc.generateExcel("IF");
            string DPath = sc.generateExcel("D");
            int day = -1;
            switch (DateTime.Now.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    day = 7;
                    break;
                case DayOfWeek.Monday:
                    day = 1;
                    break;
                case DayOfWeek.Tuesday:
                    day = 2;
                    break;
                case DayOfWeek.Wednesday:
                    day = 3;
                    break;
                case DayOfWeek.Thursday:
                    day = 4;
                    break;
                case DayOfWeek.Friday:
                    day = 5;
                    break;
                case DayOfWeek.Saturday:
                    day = 6;
                    break;
                default:
                    day = -1;
                    break;
            }
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress("FCT System Status E-KanBan@flex.com");
            mess.Subject = string.Format("FCT System Status WK{0}", Tool.WeekOfYear(DateTime.Now).ToString() + "." + day.ToString());
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
            mess.To.Add("napo.zhang@flex.com");
            mess.To.Add("jesse.he@flex.com");
            //mess.To.Add("CNSUZBU8M3TERFCTTE@flex.com");
            //mess.CC.Add("shell.shen@flex.com");
            //mess.CC.Add("robert.yu@flex.com");
            //mess.CC.Add("sarah.guo1@flex.com");
            //mess.CC.Add("Laura.Zhu@flex.com");
            //mess.CC.Add("jane.sha@flex.com");
            //mess.CC.Add("justin.mu@flex.com");
            //mess.CC.Add("jacqueline.ge@flex.com");
            Attachment uf = new Attachment(@UFPath);
            uf.Name = "UFlex Shift Report.xlsx";
            Attachment mf = new Attachment(@IFPath);
            mf.Name = "IFlex&MFlex Shift Report.xlsx";
            Attachment d = new Attachment(@DPath);
            d.Name = "Dragon Shift Report.xlsx";
            mess.Attachments.Add(uf);
            mess.Attachments.Add(mf);
            mess.Attachments.Add(d);
            string sql = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null AND (CoreBU is not null) and CoreBU <> '' ORDER BY PD, PlanShipDate, Slot";
            List<SlotPlan> list = conn.Query<SlotPlan>(sql).ToList();
            string sqlShortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE IsReceived = 0";
            List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
            List<MailSystemStatus> listStatus = new List<MailSystemStatus>();
            foreach (SlotPlan slotPlan in list)
            {
                MailSystemStatus mss = new MailSystemStatus();
                mss.Week = "WK" + slotPlan.MRP.Substring(0, slotPlan.MRP.IndexOf('.'));
                mss.Slot = slotPlan.Slot;
                mss.Type = slotPlan.Type;
                slotPlan.ShortageList = new List<SlotShortage>();
                foreach (SlotShortage item in listShortage)
                {
                    if (item.Slot == slotPlan.Slot)
                    {
                        slotPlan.ShortageList.Add(item);
                    }
                }
                mss.ShortageList = slotPlan.ShortageList;
                if (slotPlan.Pack != "" && slotPlan.Pack != null)
                {
                    mss.Station = "Pack";
                }
                else if (slotPlan.BU != "" && slotPlan.BU != null)
                {
                    mss.Station = "Button Up";
                }
                else if (slotPlan.QFAA != "" && slotPlan.QFAA != null)
                {
                    mss.Station = "QFAA";
                }
                else if (slotPlan.CSW != "" && slotPlan.CSW != null)
                {
                    mss.Station = "CSW";
                }
                else if (slotPlan.TestBU != "" && slotPlan.TestBU != null)
                {
                    mss.Station = "Test B/U";
                }
                else if (slotPlan.OI != "" && slotPlan.OI != null)
                {
                    mss.Station = "O/I";
                }
                else if (slotPlan.PV != "" && slotPlan.PV != null)
                {
                    mss.Station = "Jitter PV";
                }
                else if (slotPlan.CoreBU != "" && slotPlan.CoreBU != null)
                {
                    mss.Station = "Core B/U";
                }
                else
                {
                    mss.Station = "N/A";
                }
                listStatus.Add(mss);
            }

            string str = "<p style='color: #4183C4;'>Hi All,</p>";
            str += @"<p style='color: #4183C4;'>UltraFlex</p>
                      <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0' width='300px'>
                        <tr style='background-color: #E6F1F6; font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'  >
                          <th width='80' style='border-style: solid; border-width: 1px'>MRP</th>
                          <th width='80' style='border-style: solid; border-width: 1px'>Slot</th>
                          <th width='100' style='border-style: solid; border-width: 1px'>Station</th>
                          <th width='120' style='border-style: solid; border-width: 1px'>Shortage</th>
                        </tr>";
            foreach (MailSystemStatus item in listStatus.Where(x => x.Type == "UF").ToList())
            {
                string Shortage = "";
                foreach (SlotShortage shortage in item.ShortageList)
                {
                    Shortage = Shortage + shortage.PN + "," + shortage.QTY + ";";
                }
                str += string.Format(@"<tr style='line-height: 20px;'>
                                           <td style='border-style: solid; border-width: 1px'>{0}</td>
                                           <td style='border-style: solid; border-width: 1px'>{1}</td>
                                           <td style='border-style: solid; border-width: 1px'>{2}</td>
                                           <td style='border-style: solid; border-width: 1px'>{3}</td>
                                       </tr>", item.Week, item.Slot, item.Station, Shortage);
            }
            str += "</table>";
            str += @"<p style='color: #4183C4;'>IFlex&MFlex</p>
                      <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0' width='300px'>
                        <tr style='background-color: #E6F1F6; font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'  >
                          <th width='80' style='border-style: solid; border-width: 1px'>MRP</th>
                          <th width='80' style='border-style: solid; border-width: 1px'>Slot</th>
                          <th width='100' style='border-style: solid; border-width: 1px'>Station</th>
                          <th width='120' style='border-style: solid; border-width: 1px'>Shortage</th>
                        </tr>";
            foreach (MailSystemStatus item in listStatus.Where(x => x.Type == "IF" || x.Type == "MF").ToList())
            {
                string Shortage = "";
                foreach (SlotShortage shortage in item.ShortageList)
                {
                    Shortage = Shortage + shortage.PN + "," + shortage.QTY + ";";
                }
                str += string.Format(@"<tr style='line-height: 20px;'>
                                           <td style='border-style: solid; border-width: 1px'>{0}</td>
                                           <td style='border-style: solid; border-width: 1px'>{1}</td>
                                           <td style='border-style: solid; border-width: 1px'>{2}</td>
                                           <td style='border-style: solid; border-width: 1px'>{3}</td>
                                       </tr>", item.Week, item.Slot, item.Station, Shortage);
            }
            str += "</table>";
            str += @"<p style='color: #4183C4;'>Dragon</p>
                      <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0' width='300px'>
                        <tr style='background-color: #E6F1F6; font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'  >
                          <th width='80' style='border-style: solid; border-width: 1px'>MRP</th>
                          <th width='80' style='border-style: solid; border-width: 1px'>Slot</th>
                          <th width='100' style='border-style: solid; border-width: 1px'>Station</th>
                          <th width='120' style='border-style: solid; border-width: 1px'>Shortage</th>
                        </tr>";
            foreach (MailSystemStatus item in listStatus.Where(x => x.Type == "D").ToList())
            {
                string Shortage = "";
                foreach (SlotShortage shortage in item.ShortageList)
                {
                    Shortage = Shortage + shortage.PN + "," + shortage.QTY + ";";
                }
                str += string.Format(@"<tr style='line-height: 20px;'>
                                           <td style='border-style: solid; border-width: 1px'>{0}</td>
                                           <td style='border-style: solid; border-width: 1px'>{1}</td>
                                           <td style='border-style: solid; border-width: 1px'>{2}</td>
                                           <td style='border-style: solid; border-width: 1px'>{3}</td>
                                       </tr>", item.Week, item.Slot, item.Station, Shortage);
            }
            str += "</table>";
            str += "<p style='color: #4183C4;'>Best Regards,</p>";
            str += "<p style='color: #4183C4;'>FCT Testing</p>";
            mess.Body = str.ToString();
            client.Send(mess);
        }
    }

    internal class MailSystemStatus
    {
        public string Week { get; set; }
        public string Slot { get; set; }
        public string Type { get; set; }
        public string Station { get; set; }
        public List<SlotShortage> ShortageList { get; set; }
    }
}