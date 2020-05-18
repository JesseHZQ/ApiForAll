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
    }
}
