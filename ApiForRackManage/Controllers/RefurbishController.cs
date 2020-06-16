using ApiForFCTKB.Models;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;

namespace ApiForRackManage.Controllers
{
    public class RefurbishController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        [HttpGet]
        public void Read()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");

            Workbook wb = new Workbook(@"C:\Users\suzjeshe\Desktop\1.xls");

            List<Refurbish> rList = new List<Refurbish>();
            foreach (Worksheet ws in wb.Worksheets)
            {
                Cells cells = ws.Cells;
                DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                foreach (DataRow row in dt.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        string sn = item.ToString().Trim().ToUpper();

                        if (sn != "" && sn.Length == 7)
                        {
                            Refurbish re = new Refurbish();
                            re.SN = item.ToString().Trim().ToUpper();
                            rList.Add(re);
                        }
                    }
                    
                }
            }

            string sql = "insert into Supermarket_Refurbish values (@SN)";
            conn.Execute(sql, rList);

        }

        [HttpGet]
        public List<Refurbish> GetRefurbishes ()
        {
            string sql = "select * from Supermarket_Refurbish order by SN";
            return conn.Query<Refurbish>(sql).ToList();
        }

        [HttpPost]
        public List<Refurbish> Upload ()
        {
            List<Refurbish> list = new List<Refurbish>();
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile file = httpRequest.Files[0];
            Workbook wb = new Workbook(file.InputStream);
            Cells cells = wb.Worksheets[0].Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            List<Refurbish> listR = conn.Query<Refurbish>("select * from Supermarket_Refurbish").ToList();
            foreach (DataRow row in dt.Rows)
            {
                Refurbish refurbish = new Refurbish();
                refurbish.SN = row[0].ToString();
                if (listR.Where(x => x.SN == refurbish.SN).ToList().Count == 0)
                {
                    list.Add(refurbish);
                }
            }
            string sqlInsert = "insert into Supermarket_Refurbish (SN) VALUES (@SN)";
            conn.Execute(sqlInsert, list);
            return list;
        }
    }
}