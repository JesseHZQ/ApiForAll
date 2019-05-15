using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;
using GROUP.Manage;

namespace ApiForMinioneMiscShipping.Controllers
{
    public class BoardsController : ApiController
    {
        BaseClass bc = new BaseClass();

        [System.Web.Http.HttpPost]
        public Resp UploadDragonData(SystemList model)
        {
            foreach (var item in model.list)
            {
                DataTable dt = new DataTable();
                dt = sqlkb.ExecuteDataTable("select * from SummaryHDragon where SystemSlot = '" + item.Slot + "'");
                if (dt.Rows.Count == 0)
                {
                    sqlkb.ExecuteNonQuery("Insert into SummaryHDragon (SystemSlot, SystemModel, Customer, SO, ShipWeek, Lock, IsMore) values ('" + item.Slot + "', '" + item.Model + "','" + item.Customer + "','" + item.SO + "','" + item.MRP + "', 0, 'True')");
                }
                else
                {
                    sqlkb.ExecuteNonQuery("Update SummaryHDragon set SO = '" + item.SO + "', ShipWeek = '" + item.MRP + "' where SystemSlot = '" + item.Slot + "'");
                }
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Message = "Success";
            resp.Data = null;
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp UploadDragonCorc(CorcList model)
        {
            foreach (var item in model.list)
            {
                DataTable dt = new DataTable();
                dt = sqlkb.ExecuteDataTable("select * from SummaryHDragon where SystemSlot = '" + item.Slot + "'");
                if (dt.Rows.Count != 0)
                {
                    sqlkb.ExecuteNonQuery("Update SummaryHDragon set CORC = '" + item.Status + "', OpenIssue = '" + item.Issue + "' where SystemSlot = '" + item.Slot + "'");
                }
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Message = "Success";
            resp.Data = null;
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp UploadDragonConfig(ConfigList model)
        {
            foreach (var item in model.list)
            {
                DataTable dt = new DataTable();
                dt = sqlkb.ExecuteDataTable("select * from SummaryHDragon where SystemSlot = '" + item.Slot + "'");
                if (dt.Rows.Count != 0 && (dt.Rows[0]["SB25"] == null))
                {
                    sqlkb.ExecuteNonQuery("Update SummaryHDragon set Paradise = '" + item.Paradise + "', TeslaHC = '" + item.TeslaHC + "', SB25 = '" + item.SB25 + "', Tesla = '" + item.TeslaHD + "', Myone = '" + item.DSI + "' where SystemSlot = '" + item.Slot + "'");
                }
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Message = "Success";
            resp.Data = null;
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp get(Board board)
        {
            DataTable dt = new DataTable();
            if (board.sn == null)
            {
                dt = bc.ReadTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar]");
            }
            else
            {
                dt = bc.ReadTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN = '" + board.sn + "'");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp GetStatus()
        {
            DataTable dt = new DataTable();
            dt = bc.ReadTable("SELECT * FROM [172.21.194.214].[EKW].[dbo].[UFLEX] WHERE Line LIKE 'line%'");
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
            return ba.HttpGet("https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=qTUkez1yb2MjDemBNOaeSVXZ&client_secret=9eBwPRq9xf3ulFRXIMuCniYqeOGq6lIH", "");
        }
        /// <summary>
        /// 接收前端上传单个文件
        /// </summary>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public RespWithPath uploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile file = httpRequest.Files[0];
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);
            RespWithPath resp = new RespWithPath();
            resp.Path = Path.GetFileName(file.FileName);
            string token = httpRequest.Form["token"];
            string strbaser64 = ImgToBase64String(filePath);
            string host = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.ContentType = "application/x-www-form-urlencoded";
            request.KeepAlive = true;
            String str = "image=" + HttpUtility.UrlEncode(strbaser64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            resp.Message = result;
            return resp;
        }

        // 通用文字识别
        //public static string general()
        //{
        //    string token = "#####调用鉴权接口获取的token#####";
        //    string strbaser64 = FileUtils.getFileBase64("/work/ai/images/ocr/general.jpeg"); // 图片的base64编码
        //    string host = "https://aip.baidubce.com/rest/2.0/ocr/v1/general?access_token=" + token;
        //    Encoding encoding = Encoding.Default;
        //    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
        //    request.Method = "post";
        //    request.ContentType = "application/x-www-form-urlencoded";
        //    request.KeepAlive = true;
        //    String str = "image=" + HttpUtility.UrlEncode(strbaser64);
        //    byte[] buffer = encoding.GetBytes(str);
        //    request.ContentLength = buffer.Length;
        //    request.GetRequestStream().Write(buffer, 0, buffer.Length);
        //    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
        //    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
        //    string result = reader.ReadToEnd();
        //    Console.WriteLine("通用文字识别:");
        //    Console.WriteLine(result);
        //    return result;
        //}

        //图片 转为    base64编码的文本

        private string ImgToBase64String(string Imagefilename)
        {
            Bitmap bmp = new Bitmap(Imagefilename);
            //FileStream fs = new FileStream(Imagefilename + ".txt", FileMode.Create);
            //StreamWriter sw = new StreamWriter(fs);

            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            byte[] arr = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(arr, 0, (int)ms.Length);
            ms.Close();
            String strbaser64 = Convert.ToBase64String(arr);
            //sw.Write(strbaser64);
            //sw.Flush();
            //sw.Close();
            //fs.Dispose();
            //fs.Close();
            return strbaser64;
        }
        public class Board
        {
            public int start { get; set; }
            public int length { get; set; }
            public string sn { get; set; }

        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }

        public class RespWithPath
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
            public string Path { get; set; }
        }

        public class SystemInfo
        {
            public string Slot { get; set; }
            public string Model { get; set; }
            public string Customer { get; set; }
            public string SO { get; set; }
            public string MRP { get; set; }
            public string PD { get; set; }
            public string PO { get; set; }
        }

        public class CORC
        {
            public string Slot { get; set; }
            public string Status { get; set; }
            public string Issue { get; set; }
        }

        public class ConfigInfo
        {
            public string Slot { get; set; }
            public string Paradise { get; set; }
            public string TeslaHD { get; set; }
            public string TeslaHC { get; set; }
            public string SB25 { get; set; }
            public string DSI { get; set; }
        }

        public class CorcList
        {
            public List<CORC> list { get; set; }
        }

        public class SystemList
        {
            public List<SystemInfo> list { get; set; }
        }

        public class ConfigList
        {
            public List<ConfigInfo> list { get; set; }
        }
    }
}
