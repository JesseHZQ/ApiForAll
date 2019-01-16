using ApiForVOP.DataContext;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;

namespace ApiForAllProjects.Controllers.VOP
{

    public class VOPController : ApiController
    {
        VOPDataContext dc = new VOPDataContext();

        public class DapperService
        {
            public static SqlConnection SqlConnection()
            {
                string sqlconnectionString = ConfigurationManager.ConnectionStrings["ApiForVOP.Properties.Settings.POConnectionString"].ToString();
                var connection = new SqlConnection(sqlconnectionString);
                connection.Open();
                return connection;
            }
        }

        class POLine
        {
            public string PONumber { get; set; }
        }

        [System.Web.Http.HttpGet]
        public int addTest()
        {
            using (IDbConnection conn = DapperService.SqlConnection())
            {
                string sqlCommandText = "EXEC usp_PODetail_Extend 868733, null"; //868733, null

                var res = conn.Query(sqlCommandText).FirstOrDefault();
                var res1 = conn.Query<POLine>(sqlCommandText).FirstOrDefault();
                return 1;
            }
        }

        /// <summary>
        /// 获取首页报告 date传0时,获取所有记录  传具体年月时,获取对应的记录
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        [System.Web.Http.HttpGet]
        public Resp getReport(string date)
        {
            Resp resp = new Resp();
            if (date == "0")
            {
                List<VOP_Result> list = new List<VOP_Result>();
                try
                {
                    list = dc.VOP_Result.ToList();
                    if (list != null)
                    {
                        resp.Code = 200;
                        resp.Data = list;
                        resp.Message = "success";
                    }
                    else
                    {
                        resp.Code = 10001;
                        resp.Data = null;
                        resp.Message = "No Data";
                    }
                }
                catch (Exception ex)
                {
                    resp.Code = 10001;
                    resp.Data = null;
                    resp.Message = ex.Message;
                }
            }
            else
            {
                List<VOP_Result> list = new List<VOP_Result>();
                try
                {
                    list = dc.VOP_Result.Where(x => x.Date == date).ToList();
                    if (list.Count != 0)
                    {
                        resp.Code = 200;
                        resp.Data = list;
                        resp.Message = "success";
                    }
                    else
                    {
                        resp.Code = 10001;
                        resp.Data = null;
                        resp.Message = "No Data";
                    }
                }
                catch (Exception ex)
                {
                    resp.Code = 10001;
                    resp.Data = null;
                    resp.Message = ex.Message;
                }
            }
            return resp;
        }


        /// <summary>
        /// 添加或者修改VOP记录
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp AddOrUpdateVOP(VOP_Result model)
        {
            VOP_Result item = new VOP_Result();
            Resp resp = new Resp();
            if (dc.VOP_Result.FirstOrDefault(x => x.Date == model.Date) != null)
            {
                item = dc.VOP_Result.FirstOrDefault(x => x.Date == model.Date);
                item.Date = model.Date;
                item.Detail = model.Detail;
                item.PODetail = model.PODetail;
                item.TotalCost = model.TotalCost;
                item.PlanCost = model.PlanCost;
                item.IDMCost = model.IDMCost;
                item.TotalPercent = model.TotalPercent;
                dc.SubmitChanges();
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "更新成功";
                return resp;
            }
            else
            {
                item = new VOP_Result
                {
                    Date = model.Date,
                    Detail = model.Detail,
                    PODetail = model.PODetail,
                    TotalCost = model.TotalCost,
                    PlanCost = model.PlanCost,
                    IDMCost = model.IDMCost,
                    TotalPercent = model.TotalPercent
                };
                dc.VOP_Result.InsertOnSubmit(item);
                dc.SubmitChanges();
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "保存成功";
                return resp;
            }
        }


        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp sendMail(MailInfo model)
        {
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress(model.from);
            mess.Subject = model.subject;
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
            for (int i = 0; i < model.toList.Length; i++)
            {
                mess.To.Add(model.toList[i]);
            }
            mess.Body = model.content;
            Resp resp = new Resp();
            try
            {
                client.Send(mess);
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "success";
            }
            catch (Exception ex)
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }


        /// <summary>
        /// 接收前端上传单个文件
        /// </summary>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp uploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile file = httpRequest.Files[0];
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(file.FileName));
            Resp resp = new Resp();
            try
            {
                file.SaveAs(filePath);
                resp.Code = 200;
                resp.Data = null;
                resp.Message = filePath;
            }
            catch (Exception ex)
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }
    }

    public class Resp
    {
        public int Code { get; set; }
        public List<VOP_Result> Data { get; set; }
        public string Message { get; set; }
    }

    public class MailInfo
    {
        public string from { get; set; }
        public string subject { get; set; }
        public string[] toList { get; set; }
        public string content { get; set; }
    }
}