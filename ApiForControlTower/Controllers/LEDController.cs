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
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public List<LED> GetLEDList()
        {
            string sql = "SELECT * FROM CT_LED A LEFT JOIN [SlotKB].[dbo].[KANBAN_SLOTPLAN] B ON A.SystemName = B.Slot LEFT JOIN CT_User C ON A.OwnerId = C.UserId ORDER BY A.Item";
            return conn.Query<LED>(sql).ToList();
        }

        [HttpGet]
        public List<LED> GetLEDListByOwnerId(int Id)
        {
            string sql = "SELECT * FROM CT_LED A LEFT JOIN [SlotKB].[dbo].[KANBAN_SLOTPLAN] B ON A.SystemName = B.Slot LEFT JOIN CT_User C ON A.OwnerId = C.UserId WHERE A.OwnerId = " + Id + " ORDER BY A.Item";
            return conn.Query<LED>(sql).ToList();
        }

        [HttpGet]
        public int FinishWork(int Item)
        {
            string sql = "UPDATE CT_LED SET OwnerId = NULL WHERE Item = " + Item;
            return conn.Execute(sql);
        }

        [HttpPost]
        public int AssignWork(LED led)
        {
            string sql = "UPDATE CT_LED SET OwnerId = @OwnerId WHERE Item = @Item";
            return conn.Execute(sql, led);
        }

        [HttpGet]
        public List<EMLED> GetEMLEDList()
        {
            string sql = "SELECT * FROM [suznt004].[FCT_LED_KanBan].[dbo].[Station] WHERE ID < 6";
            return conn.Query<EMLED>(sql).ToList();
        }
    }
}
