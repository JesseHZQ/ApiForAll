using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using Dapper;

namespace ApiForVOP.Controllers
{
    public class SettingsController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ApiForVOP.Properties.Settings.DebugAssyConnectionString"].ConnectionString);

        /// <summary>
        /// 批量添加config
        /// </summary>
        [HttpPost]
        public Resp UploadSetting(Settings settings)
        {
            string sql = "INSERT INTO VOP_SETTING VALUES(@PN, @Price)";
            conn.Execute(sql, settings.list);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 单个添加config
        /// </summary>
        [HttpPost]
        public Resp addSetting(Setting setting)
        {
            string sql = "INSERT INTO VOP_SETTING VALUES(@PN, @Price)";
            conn.Execute(sql, setting);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 单个删除config
        /// </summary>
        [HttpPost]
        public Resp delSetting(Setting setting)
        {
            string sql = "DELETE FROM VOP_SETTING WHERE ID = @ID";
            conn.Execute(sql, setting);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 查询config
        /// </summary>
        [HttpGet]
        public Resp querySetting(Setting setting)
        {
            string sql = "SELECT * FROM VOP_SETTING ORDER BY PN";
            Resp resp = new Resp();
            resp.data = conn.Query<Setting>(sql).ToList();
            return resp;
        }

        /// <summary>
        /// 修改config
        /// </summary>
        [HttpPost]
        public Resp updateSetting(Setting setting)
        {
            string sql = "UPDATE VOP_SETTING SET PN = @PN, Price = @Price WHERE ID = @ID";
            conn.Execute(sql, setting);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 生成Excel
        /// </summary>
        [HttpPost]
        public Resp generateExcel(Settings settings)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook();
            wb.Worksheets.Add();
            Cells cells = wb.Worksheets[0].Cells;
            int y = 0;
            cells[y, 0].PutValue("PN");
            cells[y, 1].PutValue("Price");
            foreach (Setting list in settings.list)
            {
                y++;
                cells[y, 0].PutValue(list.PN);
                cells[y, 1].PutValue(list.Price);
            }
            string filePath = @"\\suznt004\Backup\Jesse\VOP\Files\Config_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond + ".xlsx";
            wb.Save(filePath);
            Resp resp = new Resp();
            resp.message = filePath;
            return resp;
        }

        /// <summary>
        /// 从服务器下载文件
        /// </summary>
        /// <param name="filePath"></param>
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
                return response;
            }
            catch
            {
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
        }

        public class Resp
        {
            public int code { get; set; }
            public List<Setting> data { get; set; }
            public string message { get; set; }
        }
    }

    public class Setting
    {
        public int ID { get; set; }
        public string PN { get; set; }
        public string Price { get; set; }
    }

    public class Settings
    {
        public List<Setting> list { get; set; }
    }
}
