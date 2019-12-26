using Quartz;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web;

namespace ApiForFCTKB.Job
{
    public class OnHandCheck : IJob
    {
        public string Sqlconn = "server=suznt004;uid=andy;pwd=123;database=ControlTower";
        public void Execute(IJobExecutionContext context)
        {
            DataTable dtbefore = onhandbefore();
            onhandupdate();
            DataTable dtafter = onhandafter();
            DataTable resuldt = comparedt(dtbefore, dtafter);
            if (resuldt.Rows.Count != 0)
            {
                sendmail(resuldt);
            }
        }
        public DataTable onhandbefore()
        {
            string sql = "select * from [dbo].[FA_Process_Part]";
            DataTable dt = Query(sql);
            return dt;

        }
        public void onhandupdate()
        {
            string sql = "select * from [dbo].[FA_Process_Part]";
            DataTable dt = Query(sql);
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                string fapn = dt.Rows[j]["FlexP_N"].ToString();
                string Urlstart = "http://10.194.1.19:8004/api/TdnWoApi/WoOh?itemNo=";
                string Url = Urlstart + fapn;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                request.Method = "GET";
                request.Timeout = 20000;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream myResponseStream = response.GetResponseStream();
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                string responseContent = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                response.Close();
                //123
                string oh = "";
                if (responseContent.Contains("101"))
                {
                    string[] sArray = responseContent.Split('{');
                    for (int i = 0; i < sArray.Count(); i++)
                    {
                        if (sArray[i].Contains("101"))
                        {
                            oh = sArray[i];
                            if (oh.Contains("}]"))
                            {
                                oh = oh.Replace("}]", "");
                            }
                        }
                    }
                    string[] soh = oh.Split(',');
                    string itemno = soh[1].Split(':')[1].Replace("\"", "").Trim();
                    string ohqty = soh[2].Split(':')[1].Replace("}", "").Trim();

                    string sqlupdate = "update FA_Process_Part set OH=" + ohqty + " where  FlexP_N='" + itemno + "'";
                    executeCommand(sqlupdate);
                }
            }
        }

        public DataTable onhandafter()
        {
            string sql = "select * from [dbo].[FA_Process_Part]";
            DataTable dt = Query(sql);
            return dt;
        }


        public DataTable comparedt(DataTable dtbefore, DataTable dtafter)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("LOADate", typeof(string));
            dt.Columns.Add("FlexPN", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("BuyQty", typeof(string));
            dt.Columns.Add("BeforeOH", typeof(string));
            dt.Columns.Add("AfterOH", typeof(string));

            for (int i = 0; i < dtbefore.Rows.Count; i++)
            {
                string bID = dtafter.Rows[i]["ID"].ToString();
                string bOH = (dtafter.Rows[i]["OH"] == null) ? "" : dtafter.Rows[i]["OH"].ToString();
                string OH = dtafter.Select("ID =" + bID)[0].ItemArray[9].ToString();
                if (bOH != OH)
                {
                    DataRow dr = dt.NewRow();

                    dr["LOADate"] = dtafter.Rows[i]["LOADate"].ToString();
                    dr["FlexPN"] = dtafter.Rows[i]["FlexPN"].ToString();
                    dr["Description"] = dtafter.Rows[i]["Description"].ToString();
                    dr["BuyQty"] = dtafter.Rows[i]["BuyQty"].ToString();
                    dr["BeforeOH"] = dtbefore.Rows[i]["OH"].ToString();
                    dr["AfterOH"] = dtafter.Rows[i]["OH"].ToString();
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        public int executeCommand(string sqlStr)
        {
            SqlConnection sqlConnection = new SqlConnection(Sqlconn); ;//创建数据库连接
            sqlConnection.Open();      //打开数据库连接
            SqlCommand sqlCommand = new SqlCommand(sqlStr, sqlConnection);  //执行SQL命令
            int num = sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
            return num;
        }
        public DataTable Query(string sql)
        {
            SqlConnection conn = new SqlConnection(Sqlconn);
            // 数据适配器，填充DataSet.参数1：SQL查询语句，参数2：数据库连接.
            SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
            DataSet ds = new DataSet();
            //将数据适配器中的数据填充到数据集.
            sda.Fill(ds);
            DataTable dt = ds.Tables[0];
            return dt;
        }

        public void sendmail(DataTable dt)
        {
            string to = "luna.wang@flex.com; bertie.wang@flex.com";
            MailSender mail = new MailSender();
            mail.Subject = "OnHand Change";
            mail.From = "FA SYSTEM@flex.com";
            mail.ToList = to.Split(' '); ;
            string body = "OnHand Change:";
            mail.Content = Body(body, dt);
            SendMail(mail);
        }
        public class MailSender
        {
            public string Subject { get; set; }
            public string From { get; set; }
            public string[] ToList { get; set; }
            public string[] CCList { get; set; }
            public string[] FilePathList { get; set; }
            public string Content { get; set; }

        }

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

        public class MailResp
        {
            public string Code { get; set; }
            public string Message { get; set; }
        }


        public string Body(string str, DataTable dt)
        {
            string strbody = "";
            strbody = "<p style=\"font-size: 10pt\">" + str + "</p><table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">";
            strbody += "<div align=\"center\">";
            strbody += "<tr>";
            for (int hcol = 0; hcol < dt.Columns.Count; hcol++)
            {
                strbody += "<td bgcolor=\"E0E0E0\">&nbsp;&nbsp;&nbsp;";
                strbody += dt.Columns[hcol].ColumnName;
                strbody += "&nbsp;&nbsp;&nbsp;</td>";
            }
            strbody += "</tr>";
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                strbody += "<tr>";
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    strbody += "<td bgcolor=\"ffffff\">&nbsp;&nbsp;&nbsp;";
                    strbody += dt.Rows[row][col].ToString();
                    strbody += "&nbsp;&nbsp;&nbsp;</td>";
                }
                strbody += "</tr>";
            }
            strbody += "</table>";
            strbody += "</div>";
            return strbody;
        }
    }
}