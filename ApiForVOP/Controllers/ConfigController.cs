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
    public class ConfigController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ApiForVOP.Properties.Settings.DebugAssyConnectionString"].ConnectionString);

        /// <summary>
        /// 批量添加config
        /// </summary>
        [HttpPost]
        public Resp UploadConfig(Configs configs)
        {
            string sql = "INSERT INTO VOP_CONFIG VALUES(@Parent, @BOM, @QTY)";
            conn.Execute(sql, configs.list);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 单个添加config
        /// </summary>
        [HttpPost]
        public Resp addConfig(Config config)
        {
            string sql = "INSERT INTO VOP_CONFIG VALUES(@Parent, @BOM, @QTY)";
            conn.Execute(sql, config);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 单个删除config
        /// </summary>
        [HttpPost]
        public Resp delConfig(Config config)
        {
            string sql = "DELETE FROM VOP_CONFIG WHERE ID = @ID";
            conn.Execute(sql, config);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 查询config
        /// </summary>
        [HttpGet]
        public Resp queryConfig(Config config)
        {
            string sql = "SELECT * FROM VOP_CONFIG ORDER BY PARENT, BOM";
            Resp resp = new Resp();
            resp.data = conn.Query<Config>(sql).ToList();
            return resp;
        }

        /// <summary>
        /// 修改config
        /// </summary>
        [HttpPost]
        public Resp updateConfig(Config config)
        {
            string sql = "UPDATE VOP_CONFIG SET BOM = @BOM, QTY = @QTY WHERE ID = @ID";
            conn.Execute(sql, config);
            Resp resp = new Resp();
            return resp;
        }

        /// <summary>
        /// 生成Excel
        /// </summary>
        [HttpPost]
        public Resp generateExcel(Configs configs)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook();
            wb.Worksheets.Add();
            Cells cells = wb.Worksheets[0].Cells;
            int y = 0;
            cells[y, 0].PutValue("PCBA/System");
            cells[y, 1].PutValue("BOM PN");
            cells[y, 2].PutValue("QTY");
            foreach (Config list in configs.list)
            {
                y++;
                cells[y, 0].PutValue(list.Parent);
                cells[y, 1].PutValue(list.BOM);
                cells[y, 2].PutValue(list.QTY);
            }
            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\VOP\Files\BOM_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond + ".xlsx";
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
    }

    public class Config
    {
        public int ID { get; set; }
        public string Parent { get; set; }
        public string BOM { get; set; }
        public string QTY { get; set; }
    }

    public class Configs
    {
        public List<Config> list { get; set; }
    }

    public class Resp
    {
        public int code { get; set; }
        public List<Config> data { get; set; }
        public string message { get; set; }
    }
}
