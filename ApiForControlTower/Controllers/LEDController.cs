using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ApiForControlTower.Models;
using Dapper;

namespace ApiForControlTower.Controllers
{
    public class LEDController : ApiController
    {
        public IDbConnection conn004 = new SqlConnection(ConfigurationManager.ConnectionStrings["suznt004"].ConnectionString);

        [HttpGet]
        public List<LED> GetLEDList()
        {
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[CT_LED] A LEFT JOIN [dbo].[KANBAN_SLOTPLAN] B ON A.SystemName = B.Slot collate Chinese_PRC_CI_AS ORDER BY A.Item";
            return conn004.Query<LED>(sql).ToList();
        }

        [HttpGet]
        public List<EMLED> GetEMLEDList()
        {
            string sql = "SELECT * FROM [suznt004].[FCT_LED_KanBan].[dbo].[Station] WHERE ID < 6";
            return conn004.Query<EMLED>(sql).ToList();
        }
    }
}
