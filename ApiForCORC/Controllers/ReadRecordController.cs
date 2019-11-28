using ApiForCORC.Models;
using Aspose.Cells;
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

namespace ApiForCORC.Controllers
{
    public class ReadRecordController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);
        public int GetSOInfo()
        {
            List<CheckInInfo> list = new List<CheckInInfo>();
            LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);
            Workbook wb = new Workbook(@"C:\Users\suzjeshe\Desktop\最后5000.csv", loadOptions);
            Cells cells = wb.Worksheets[0].Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, 500000, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                CheckInInfo info = new CheckInInfo();
                info.Name = item["Name"].ToString();
                info.CtfTp = item["CtfTp"].ToString();
                info.CtfId = item["CtfId"].ToString();
                info.Gender = item["Gender"].ToString();
                info.Birthday = item["Birthday"].ToString();
                info.Address = item["Address"].ToString();
                info.Mobile = item["Mobile"].ToString();
                info.Date = item["Version"].ToString();
                list.Add(info);
            }
            string sql = "insert into CT_CheckInRecord (Name, CtfTp, CtfId, Gender, Birthday, Address, Mobile, Date) values (@Name, @CtfTp, @CtfId, @Gender, @Birthday, @Address, @Mobile, @Date)";
            return conn.Execute(sql, list);
        }
    }
}