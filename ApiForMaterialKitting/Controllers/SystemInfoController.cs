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

namespace ApiForMaterialKitting.Controllers
{
    public class SystemInfoController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp getAllSystem()
        {
            DataTable dt = new DataTable();
            int wknow = GetWeekOfYear(DateTime.Now);
            int wk = GetWeekOfYear(DateTime.Now.AddDays(14));
            if (wk > wknow)
            {
                dt = SqlHelper.ExecuteDataTable("select a.*,b.*  from(SELECT * FROM SummaryH where Lock = 0 and ShipWeek <= " + wk + " and ShipWeek >= " + wknow + ")a left join (select * from MaterialKitting where IsDel = 0)b on a.SystemSlot = b.Slot order by a.ShipWeek");
            } 
            else
            {
                dt = SqlHelper.ExecuteDataTable("select a.*,b.*  from(SELECT * FROM SummaryH where Lock = 0 and ShipWeek <= " + wk + " or ShipWeek >= " + wknow + ")a left join (select * from MaterialKitting where IsDel = 0)b on a.SystemSlot = b.Slot order by a.ShipWeek");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public string getBaiduToken()
        {
            baiduAI ba = new baiduAI();
            return ba.HttpGet("https://openapi.baidu.com/oauth/2.0/token?grant_type=client_credentials&client_id=tUEDZ8hkSL1IaUb8kv8CUNR7&client_secret=2Fq6ai1O7GP3ZMygplwNmExwEIKs4Yjv", "");
        }

        [System.Web.Http.HttpGet]
        public Resp getTypedSystem(string type)
        {
            DataTable dt = new DataTable();
            int wk = GetWeekOfYear(DateTime.Now) + 2;
            if (type == "all")
            {
                dt = SqlHelper.ExecuteDataTable("select a.*,b.*  from(SELECT * FROM MaterialKitting where IsDel = 0)a left join SummaryH b on b.SystemSlot = a.Slot order by DemandDate");
            }
            else
            {
                dt = SqlHelper.ExecuteDataTable("select a.*,b.*  from(SELECT * FROM MaterialKitting where IsDel = 0)a left join SummaryH b on b.SystemSlot = a.Slot where a.station = '" + type + "' order by DemandDate");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getSystemById(int id)
        {
            DataTable dt = new DataTable();
            int wk = GetWeekOfYear(DateTime.Now) + 2;
            dt = SqlHelper.ExecuteDataTable("select a.*,b.*  from(SELECT * FROM MaterialKitting where IsDel = 0)a left join SummaryH b on b.SystemSlot = a.Slot where a.id = '" + id + "'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updateKitting(KittingModelId model)
        {
            Resp resp = new Resp();
            string timestampStr = model.RequireDate;
            long timestamp = long.Parse(timestampStr);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0).AddMilliseconds(timestamp);
            int index = SqlHelper.ExecuteNonQuery("update MaterialKitting set Station = '" + model.Station + "', Remark = '" + model.Remark + "', DemandDate = '" + dt.ToLocalTime() + "' where id = " + model.Id);
            if (index != 0)
            {
                resp.Code = 200;
                resp.Message = "修改成功";
                return resp;
            }
            resp.Code = 10001;
            resp.Message = "修改失败";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp beginTask(int id)
        {
            int count = SqlHelper.ExecuteNonQuery("update MaterialKitting set StartDate = '" + DateTime.Now + "' where id = '" + id + "'");
            Resp resp = new Resp();
            if (count > 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "执行成功";
            }
            else
            {

                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "执行失败";
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp finishTask(int id)
        {
            int count = SqlHelper.ExecuteNonQuery("update MaterialKitting set FinishDate = '" + DateTime.Now + "' where id = '" + id + "'");
            Resp resp = new Resp();
            if (count > 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "执行成功";
            }
            else
            {

                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "执行失败";
            }
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp finishAssign(MultipleId model)
        {
            foreach (var item in model.arr)
            {
                SqlHelper.ExecuteNonQuery("update MaterialKitting set AssignDate = '" + DateTime.Now + "' where id = '" + item + "'");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = null;
            resp.Message = "执行成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getPickList(string po)
        {
            DataTable dt = new DataTable();
            dt = SqlHelperTOV.ExecuteDataTable("EXEC usp_PODetail_Extend " + po + ", null");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updatePickList(PickList model)
        {
            Resp resp = new Resp();
            int index = SqlHelper.ExecuteNonQuery("update MaterialKitting set PickList = '" + model.str + "' where Id = '" + model.id + "' ");
            if (index != 0)
            {
                resp.Code = 200;
                resp.Message = "更新成功";
                return resp;
            }
            resp.Code = 10001;
            resp.Message = "更新失败";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp sendBeginEmail(KittingModel model)
        {
            Resp resp = new Resp();
            DataTable d = new DataTable();
            string requester = "Admin";
            d = SqlHelper.ExecuteDataTable("select Name from MaterialKittingUser where Emp = '" + model.RequestorId + "'" );
            if (d.Rows.Count > 0)
            {
                requester = d.Rows[0]["Name"].ToString();
            }
            string timestampStr = model.RequireDate;
            long timestamp = long.Parse(timestampStr);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0).AddMilliseconds(timestamp);

            string str = @"<table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0'><thead>";
            str += "<tr style='font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'>";
            str += "<th width='80' style='border-style: solid; border-width: 1px'>Slot</th>";
            str += "<th width='80' style='border-style: solid; border-width: 1px'>Type</th>";
            str += "<th width='120' style='border-style: solid; border-width: 1px'>Customer</th>";
            str += "<th width='100' style='border-style: solid; border-width: 1px'>PO</th>";
            str += "<th width='80' style='border-style: solid; border-width: 1px'>Station</th>";
            str += "<th width='120' style='border-style: solid; border-width: 1px'>Demand time</th>";
            str += "<th width='100' style='border-style: solid; border-width: 1px'>Requester</th></tr>";
            str += "</thead><tbody>";
            str += "<tr style='font-size: 14px; height: 24px; text-align: center;'>";
            str += "<td style='border-style: solid; border-width: 1px'>" + model.SystemSlot + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + "UF" + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + "Teradyne" + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + model.PO + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + model.Station.ToUpper() + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + dt.ToLocalTime() + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + requester + "</td>";
            str += "</tr>";
            str += "</tbody></table>";
            str += "<a href='http://suznt004:8082/materialkitting'>详情请点击跳转，建议使用chrome浏览器</a>";

            int index = SqlHelper.ExecuteNonQuery("insert into MaterialKitting (Slot, PO, Station, Remark, DemandDate, InformDate, IsDel) values ('" + model.SystemSlot + "','" + model.PO + "','" + model.Station + "','" + model.Remark + "','" + dt.ToLocalTime() + "','" + DateTime.Now + "', 0)");
            if (index != 0)
            {
                string[] list = new string[] {
                    "sky.zong@flex.com",
                    "xiaohong.shi@flex.com"
                };
                string[] cclist = new string[]
                {
                    "robert.yu@flex.com",
                    "napo.zhang@flex.com",
                    "gary.huang@flex.com",
                    "yu.you@flex.com",
                    "hongliang.xu1@flex.com",
                    "yan.yu@flex.com",
                    "sean.hu@flex.com",
                    "evan.luo@flex.com",
                    "yufei.huang@flex.com",
                    "jie.li@flex.com",
                    "jane.zhang1@flex.com",
                    "rachel.wang2@flex.com",
                    "zhihong.zhang@flex.com",
                    "lisa.qian@flex.com",
                    "chris.liu1@flex.com",
                    "eric.lin2@flex.com",
                    "Jesse.He@flex.com"
                };
                int result = SendMail("Material-Kitting@flex.com", "Ultra system 增料", list, cclist, str);
                if (result == 1)
                {
                    resp.Code = 200;
                    resp.Message = "发送成功";
                }
                if (result == 0)
                {
                    resp.Code = 10001;
                    resp.Message = "发送失败";
                }
                return resp;
            }
            resp.Code = 10001;
            resp.Message = "发送失败";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp sendDeleteEmail(int id, string slot, string station)
        {
            Resp resp = new Resp();
            int index = SqlHelper.ExecuteNonQuery("update MaterialKitting set IsDel = 1 where id =" + id);
            if (index != 0)
            {
                resp.Code = 200;
                resp.Message = "备料撤销成功";
                return resp;
            }
            resp.Code = 10001;
            resp.Message = "备料撤销失败";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp sendEnsureEmail(int id, string slot, string station)
        {
            Resp resp = new Resp();
            int index = SqlHelper.ExecuteNonQuery("update MaterialKitting set EnsureDate = '" + DateTime.Now + "' where id =" + id);
            if (index != 0)
            {
                resp.Code = 200;
                resp.Message = "确认收料成功";
                return resp;
            }
            resp.Code = 10001;
            resp.Message = "确认收料失败";
            return resp;
        }
        
        /// <summary>
        /// 获取指定日期，在为一年中为第几周
        /// </summary>
        /// <param name="dt">指定时间</param>
        /// <reutrn>返回第几周</reutrn>
        private static int GetWeekOfYear(DateTime dt)
        {
            GregorianCalendar gc = new GregorianCalendar();
            int weekOfYear = gc.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            return weekOfYear;
        }

        /// <summary>
        /// 发送邮件的统一方法
        /// </summary>
        /// <param name="from"></param>
        /// <param name="subject"></param>
        /// <param name="nameList"></param>
        /// <param name="content"></param>
        public static int SendMail(string from, string subject, string[] nameList, string[] nameCCList, string content)
        {
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress(from);
            mess.Subject = string.Format(subject);
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
            foreach (string item in nameList)
            {
                mess.To.Add(item);
            }
            foreach (string item in nameCCList)
            {
                mess.CC.Add(item);
            }
            mess.Body = string.Format(content);
            try
            {
                client.Send(mess);
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public class KittingModel
        {
            public string SystemSlot { get; set; }
            public string PO { get; set; }
            public string Station { get; set; }
            public string RequireDate { get; set; }
            public string RequestorId { get; set; }
            public string Remark { get; set; }

        }
        public class KittingModelId
        {
            public int Id { get; set; }
            public string SystemSlot { get; set; }
            public string PO { get; set; }
            public string Station { get; set; }
            public string RequireDate { get; set; }
            public string Remark { get; set; }

        }


        public class PickList
        {
            public int id { get; set; }
            public string str { get; set; }
        }

        public class MultipleId
        {
            public int[] arr { get; set; }
        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }
    }
}