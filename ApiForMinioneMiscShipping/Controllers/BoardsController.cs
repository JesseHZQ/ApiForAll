﻿using System;
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
            file.SaveAs(filePath);
            Resp resp = new Resp();
            string token = "24.af28dd056f4ec0d136df87cf5b266037.2592000.1545982236.282335-14817751";
            string strbaser64 = ImgToBase64String(filePath);
            string host = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=" + token;
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
            FileStream fs = new FileStream(Imagefilename + ".txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            byte[] arr = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(arr, 0, (int)ms.Length);
            ms.Close();
            String strbaser64 = Convert.ToBase64String(arr);
            sw.Write(strbaser64);
            sw.Flush();
            sw.Close();
            fs.Dispose();
            fs.Close();
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
    }
}
