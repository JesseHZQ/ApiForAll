using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ApiForCommonTools.Controllers
{
    public class FileController : ApiController
    {
        /// <summary>
        /// 从服务器下载文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fileName">自定义文件名</param>
        /// <returns></returns>
        [HttpGet]
        public HttpResponseMessage downloadFile(string filePath, string fileName)
        {
            try
            {
                var stream = new FileStream(filePath, FileMode.Open);
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = fileName
                };
                //stream.Dispose();
                //stream.Close();
                return response;
            }
            catch (Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
        }

        /// <summary>
        /// 接收前端上传单个文件
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public string uploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile file = httpRequest.Files[0];
            if (!Directory.Exists(HttpContext.Current.Server.MapPath("../../Uploads")))
            {
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("../../Uploads"));
            }
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);
            return filePath;
        }
    }
}