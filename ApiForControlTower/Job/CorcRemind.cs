using ApiForControlTower.Controllers;
using Quartz;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Web;

namespace ApiForControlTower.Job
{
    public class CorcRemind : IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            CORCController cc = new CORCController();
            List<CORC> cList = cc.GetCORC();
            cList = cList.Where(x => (DateTime.Now - x.InsertTime).Days >= 3 && x.CloseTime.Year == 1).ToList();
            MailSender ms = new MailSender();
            string date = DateTime.Now.ToString("dddd, MMM d", DateTimeFormatInfo.InvariantInfo);
            ms.Subject = string.Format("CORC Alert on {0}", date);
            ms.From = "Control Tower Control.Tower@flex.com";
            ms.ToList = new List<string>();
            //ms.ToList.Add("aivin.shao@flex.com");
            //ms.ToList.Add("napo.zhang@flex.com");
            ms.ToList.Add("jesse.he@flex.com");
            ms.CCList = new List<string>();
            ms.FilePathList = new List<string>();
            ms.Content = GetMailContent(cList);
            SendMail(ms);
        }

        public string GetMailContent(List<CORC> l)
        {
            string content = string.Format(@"<p style='color: #000; margin-bottom: 20px;'>Hello Everyone,</p>
              <p style='color: #000; margin-bottom: 20px;'>Here is CORC status alert on {0}.</p>
              <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #000;' cellspacing='0'>
                <tr style='background-color: #5294f7; font-weight: bold; color: #000; font-size: 14px; height: 24px; text-align: center;'  >
                  <th width='100' style='border-style: solid; border-width: 1px'>Slot</th>
                  <th width='100' style='border-style: solid; border-width: 1px'>PD Week</th>
                  <th width='250' style='border-style: solid; border-width: 1px'>Open Date</th>
                  <th width='120' style='border-style: solid; border-width: 1px'>Aging Days</th>
                  <th width='200' style='border-style: solid; border-width: 1px'>Issue Owner</th>
                  <th width='100' style='border-style: solid; border-width: 1px'>CORC Issue</th>
                </tr>", DateTime.Now.ToString("MMM d", DateTimeFormatInfo.InvariantInfo));
            foreach (CORC item in l)
            {
                string single = string.Format(@"<tr style='line-height: 20px;'>
                    <td style='border-style: solid; border-width: 1px'>{0}</td>
                    <td style='border-style: solid; border-width: 1px'>{1}</td>
                    <td style='border-style: solid; border-width: 1px'>{2}</td>
                    <td style='border-style: solid; border-width: 1px'>{3}</td>
                    <td style='border-style: solid; border-width: 1px'>{4}</td>
                    <td style='border-style: solid; border-width: 1px; text-align: left;'>{5}</td>
                    </tr>", item.Slot, item.PD, item.InsertTime, (DateTime.Now - item.InsertTime).Days, "Napo Zhang", item.CORC_Issue);
                content += single;
            }
            content += "</table>";
            content += "<p style='color: #000; font-weight: 700;'>Note: </p>";
            content += "<p style='color: #000;'>CORC Issue's aging days >= 3 will be alarmed by this mail!</p>";
            content += "<div style='margin: 20px 0 0;'>Best Regards!</div>";
            content += "<div style='margin: 0 0 50px;'>CORC Alert System</div>";
            content += "<p style='font-style: italic;'>Auto mail, please do not reply, thanks!</p>";
            return content;
        }

        public string SendMail(MailSender mailSender)
        {
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress(mailSender.From);
            mess.Subject = string.Format(mailSender.Subject);
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("root", "Flex_1969");
            foreach (string item in mailSender.ToList)
            {
                mess.To.Add(item);
            }
            if (mailSender.CCList != null)
            {
                foreach (string item in mailSender.CCList)
                {
                    mess.CC.Add(item);
                }
            }
            if (mailSender.FilePathList != null)
            {
                foreach (string item in mailSender.FilePathList)
                {
                    Attachment at = new Attachment(@item, MediaTypeNames.Application.Octet);
                    //at.Name = item.Substring(item.LastIndexOf('\\'));
                    at.Name = "CORC_Alert.xlsx";
                    mess.Attachments.Add(at);
                }
            }
            mess.Body = string.Format(mailSender.Content);
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

        public class MailSender
        {
            public string Subject { get; set; }
            public string From { get; set; }
            public List<string> ToList { get; set; }
            public List<string> CCList { get; set; }
            public List<string> FilePathList { get; set; }
            public string Content { get; set; }
        }
    }
}