using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Web;
using ApiForFCTKBNew.Controllers;
using ApiForFCTKBNew.Models;
using Dapper;
using Quartz;
using Quartz.Impl;

namespace ApiForFCTKBNew.Job
{
    public class Mail : IJob
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        public void Execute(IJobExecutionContext context)
        {
            SendMail();
        }

        public string GetMailContent()
        {
            string content = @"<p style='color: #4183C4;'>Following is the daily issue for FCT System, please take actions for these issues, thank you!</p>
              <p style='color: #4183C4;'>For details please refer to <span style='font-weight: 700; font-style: italic;'>http://10.194.51.14:8086/syskb</span> <span style='font-weight: 700;'> (Please use Google Chrome) </span> </p>
              <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0' width='1101px'>
                <tr style='background-color: #E6F1F6; font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'  >
                  <th width='60' style='border-style: solid; border-width: 1px'>ShipDay</th>
                  <th width='60' style='border-style: solid; border-width: 1px'>Type</th>
                  <th width='80' style='border-style: solid; border-width: 1px'>Slot</th>
                  <th width='120' style='border-style: solid; border-width: 1px'>Customer</th>
                  <th width='80' style='border-style: solid; border-width: 1px'>MRP Week</th>
                  <th width='230' style='border-style: solid; border-width: 1px'>CORC Issue</th>
                  <th width='230' style='border-style: solid; border-width: 1px'>Engineering Issue</th>
                  <th width='230' style='border-style: solid; border-width: 1px'>Shortage</th>
                </tr>";

            string sql = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null ORDER BY MRP";
            List<SlotPlan> list = conn.Query<SlotPlan>(sql).ToList().Where(x => Tool.GetSlotPlanValidate(x.MRP) == true).ToList();
            string sqlConfig = "SELECT * FROM KANBAN_SLOTCONFIG";
            List<SlotConfig> listConfig = conn.Query<SlotConfig>(sqlConfig).ToList();
            string sqlShortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE IsReceived = 0";
            List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
            string sqlEIssue = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Status <> 'Close'";
            List<SlotEIssue> listEIssue = conn.Query<SlotEIssue>(sqlEIssue).ToList();
            foreach (SlotPlan slotplan in list)
            {
                slotplan.ConfigList = new List<SlotConfig>();
                slotplan.ShortageList = new List<SlotShortage>();
                slotplan.EIssueList = new List<SlotEIssue>();
                foreach (SlotConfig item in listConfig)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ConfigList.Add(item);
                    }
                }
                foreach (SlotShortage item in listShortage)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ShortageList.Add(item);
                    }
                }
                foreach (SlotEIssue item in listEIssue)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.EIssueList.Add(item);
                    }
                }
            }
            List<SlotPlan> spList = list.Where(x => (x.CORC_Issue != "" && x.CORC_Issue != null) || (x.ShortageList.Count > 0) || (x.EIssueList.Count > 0)).ToList();
            foreach (SlotPlan sp in spList)
            {
                string EIssue = "";
                foreach (SlotEIssue item in sp.EIssueList)
                {
                    EIssue += item.Item;
                }
                string Shortage = "";
                foreach (SlotShortage item in sp.ShortageList)
                {
                    Shortage = Shortage + item.PN + "," + item.QTY + ";";
                }
                string single = string.Format(@"<tr style='line-height: 20px;'>
                    <td style='border-style: solid; border-width: 1px'>{0}</td>
                    <td style='border-style: solid; border-width: 1px'>{1}</td>
                    <td style='border-style: solid; border-width: 1px'>{2}</td>
                    <td style='border-style: solid; border-width: 1px'>{3}</td>
                    <td style='border-style: solid; border-width: 1px'>{4}</td>
                    <td style='border-style: solid; border-width: 1px; text-align: left;'>{5}</td>
                    <td style='border-style: solid; border-width: 1px; text-align: left;'>{6}</td>
                    <td style='border-style: solid; border-width: 1px; text-align: left;'>{7}</td>
                    </tr>", sp.PlanShipDate, sp.Type == "D" ? "Dragon" : sp.Type, sp.Slot, sp.Customer, sp.MRP, sp.CORC_Issue, EIssue, Shortage);
                content += single;
            }
            content += "</table>";
            return content;
        }

        public void SendMail()
        {
            string x = DateTime.Now.ToString("dddd, MMM d", DateTimeFormatInfo.InvariantInfo);
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress("FCT System Status E-KanBan@flex.com");
            mess.Subject = string.Format("FCT System Status E-KanBan Daily Report on {0}", x);
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
            mess.To.Add("shell.shen@flex.com");
            mess.To.Add("robert.yu@flex.com");
            mess.To.Add("napo.zhang@flex.com");
            mess.To.Add("city.jiang@flex.com");
            mess.To.Add("eric.lin2@flex.com");
            mess.To.Add("justin.mu@flex.com");
            mess.To.Add("jacqueline.ge@flex.com");
            mess.To.Add("ivy.chen2@flex.com");
            mess.To.Add("shirley.li2@flex.com");
            mess.To.Add("judy.ye@flex.com");
            mess.To.Add("jet.wang@flex.com");
            mess.To.Add("mike.fu@flex.com");
            mess.To.Add("sarah.guo1@flex.com");
            mess.To.Add("ya.gu@flex.com");
            mess.To.Add("jane.sha@flex.com");
            mess.To.Add("justin.huang@flex.com");
            mess.To.Add("aivin.shao@flex.com");
            mess.To.Add("brandon.gan@flex.com");
            mess.To.Add("evan.luo@flex.com");
            mess.To.Add("tina.hu@flex.com");
            mess.To.Add("martin.hao@flex.com");
            mess.To.Add("tony.gu@flex.com");
            mess.To.Add("zhuqi.xu@flex.com");
            mess.To.Add("CNSUZBU8M3TERFCTTE@flex.com");
            string str = GetMailContent();
            mess.Body = str.ToString();
            client.Send(mess);
        }
    }
}