using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;
using Aspose.Cells;
using System.Web;

namespace ApiForCORC.Controllers
{
    public class THController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["TH"].ConnectionString);

        [HttpGet]
        public void GetTHListExcel(string from, string to)
        {
            from = from + " 00:00:00";
            to = to + " 23:59:59";
            string sql = "select * from EDW_TEMPERATURE_HUMIDITY where CreationTime >= '" + from + "' and CreationTime <= '" + to + "'";
            List<TH> thList = conn.Query<TH>(sql).ToList();

            // 注册Aspose License
            License license = new License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;
            int y = 0;
            cells[y, 0].PutValue("DateTime");
            cells[y, 1].PutValue("Location");
            cells[y, 2].PutValue("Temperature");
            cells[y, 3].PutValue("Humidity");

            foreach (TH th in thList)
            {
                y++;
                cells[y, 0].PutValue(th.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"));
                cells[y, 1].PutValue(th.Location);
                cells[y, 2].PutValue(th.Temperature);
                cells[y, 3].PutValue(th.Humidity);
            }
            wb.Save(HttpContext.Current.Response, "Temperature&Humidity.xlsx", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Xlsx));
        }



        public class TH
        {
            public int ID { get; set; }
            public DateTime CreationTime { get; set; }
            public decimal Temperature { get; set; }
            public decimal Humidity { get; set; }
            public string Location { get; set; }
        }
    }
}
