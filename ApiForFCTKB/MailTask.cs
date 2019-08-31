using ApiForFCTKB.Controllers;
using ApiForFCTKB.Models;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ApiForFCTKB
{
    public class MailTask
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
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
                  <th width='40' style='border-style: solid; border-width: 1px'>PD</th>
                  <th width='230' style='border-style: solid; border-width: 1px'>CORC Issue</th>
                  <th width='230' style='border-style: solid; border-width: 1px'>Engineering Issue</th>
                </tr>";
            string sql = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTPLAN] WHERE ShippingDate is null ORDER BY PD";
            List<SlotPlan> spList = conn.Query<SlotPlan>(sql).ToList().Where(x => GetValidate(x.MRP) == true && ((x.CORC_Issue != "" && x.CORC_Issue != null) || (x.Engineering_Issue != "" && x.Engineering_Issue != "[]" && x.Engineering_Issue != null))).ToList();
            foreach (SlotPlan sp in spList)
            {
                if (sp.Engineering_Issue != "" && sp.Engineering_Issue != null)
                {
                    JArray jArray = (JArray)JsonConvert.DeserializeObject(sp.Engineering_Issue);
                    string E_Issue = "";
                    if (jArray.Count > 0)
                    {
                        foreach (var item in jArray)
                        {
                            if (((JObject)item)["Status"].ToString() == "Open")
                            {
                                string a = ((JObject)item)["Item"].ToString();
                                E_Issue += a;
                            }
                        }
                    }
                    string single = string.Format(@"<tr style='line-height: 20px;'>
                        <td style='border-style: solid; border-width: 1px'>{0}</td>
                        <td style='border-style: solid; border-width: 1px'>{1}</td>
                        <td style='border-style: solid; border-width: 1px'>{2}</td>
                        <td style='border-style: solid; border-width: 1px'>{3}</td>
                        <td style='border-style: solid; border-width: 1px'>{4}</td>
                        <td style='border-style: solid; border-width: 1px; text-align: left;'>{5}</td>
                        <td style='border-style: solid; border-width: 1px; text-align: left;'>{6}</td>
                      </tr>", sp.PlanShipDate, sp.Type == "D" ? "Dragon" : sp.Type, sp.Slot, sp.Customer, sp.PD, sp.CORC_Issue, E_Issue);
                    content += single;
                }
                else
                {
                    string single = string.Format(@"<tr style='line-height: 20px;'>
                        <td style='border-style: solid; border-width: 1px'>{0}</td>
                        <td style='border-style: solid; border-width: 1px'>{1}</td>
                        <td style='border-style: solid; border-width: 1px'>{2}</td>
                        <td style='border-style: solid; border-width: 1px'>{3}</td>
                        <td style='border-style: solid; border-width: 1px'>{4}</td>
                        <td style='border-style: solid; border-width: 1px; text-align: left;'>{5}</td>
                        <td style='border-style: solid; border-width: 1px; text-align: left;'>{6}</td>
                      </tr>", sp.PlanShipDate, sp.Type == "D" ? "Dragon" : sp.Type, sp.Slot, sp.Customer, sp.PD, sp.CORC_Issue, sp.Engineering_Issue);
                    content += single;
                }
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
        public bool GetValidate(string mrpStr)
        {
            int week = Tool.WeekOfYear(DateTime.Now); // 获取上周周别
            float mrp = float.Parse(mrpStr); // 获取MRP
            int range = 5; // 抓取的周范围 后期改成可修改
            bool IsRange = true; // 定义周范围的Bool
            if (week + range <= 54) // 年底之前逻辑简单
            {
                IsRange = (mrp >= week && mrp <= week + range);
            }
            else // 年底的时候周别逻辑
            {
                IsRange = (mrp >= week && mrp <= 54) || (mrp >= 1 && mrp <= week + range - 53);
            }
            return IsRange;
        }
    }
}