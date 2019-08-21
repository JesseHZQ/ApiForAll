using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Net.Mime;
using System.Web.Http;

namespace ApiForControlTower.Controllers
{
    public partial class MailController : ApiController
    {
        /// <summary>
        /// 发送邮件的接口 可以添加附件
        /// </summary>
        /// <param name="mailSender"></param>
        /// <returns></returns>
        [HttpPost]
        public MailResp SendMail(MailSender mailSender)
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
                    at.Name = item.Substring(item.LastIndexOf('/'));
                    mess.Attachments.Add(at);
                }
            }
            mess.Body = string.Format(mailSender.Content);
            MailResp resp = new MailResp();
            try
            {
                client.Send(mess);
                resp.Code = "200";
                resp.Message = "Success";
            }
            catch (Exception ex)
            {
                resp.Code = "10001";
                resp.Message = "Failed! " + ex.Message;
            }
            return resp;
        }
    }
}
