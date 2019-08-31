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
        public IDbConnection connEM = new SqlConnection(ConfigurationManager.ConnectionStrings["EMLED"].ConnectionString);

        [HttpGet]
        public List<LED> GetLEDList()
        {
            string sql = "SELECT * FROM CT_LEDMaster A " +
                         "LEFT JOIN [SlotKB].[dbo].[KANBAN_SLOTPLAN] B " +
                         "ON A.SystemName = B.Slot " +
                         "LEFT JOIN CT_User C " +
                         "ON A.OwnerId = C.UserId " +
                         "ORDER BY A.Item";
            return conn.Query<LED>(sql).ToList();
        }

        [HttpGet]
        public List<LED> GetLEDListByOwnerId(int Id)
        {
            string sql = "SELECT * FROM CT_LEDMaster WHERE OwnerId = " + Id + " ORDER BY Item";
            return conn.Query<LED>(sql).ToList();
        }

        [HttpGet]
        public int FinishWork(int Item)
        {
            string sql = "UPDATE CT_LEDMaster SET OwnerId = NULL WHERE Item = " + Item;
            return conn.Execute(sql);
        }

        [HttpPost]
        public int AssignWork(LED led)
        {
            string sql = "UPDATE CT_LEDMaster SET OwnerId = @OwnerId WHERE Item = @Item";
            return conn.Execute(sql, led);
        }

        [HttpGet]
        public List<EMLED> GetEMLEDList()
        {
            string sql = "SELECT * FROM [suznt004].[FCT_LED_KanBan].[dbo].[Station] WHERE ID < 6 OR ID IN (9, 13)";
            return conn.Query<EMLED>(sql).ToList();
        }

        [HttpGet]
        public List<EMLEDSteps> GetEMLEDStepsByLogID(int LogID)
        {
            string sql = "SELECT * FROM Steplog WHERE Aslog = " + LogID + " and EndTime is not null";
            return connEM.Query<EMLEDSteps>(sql).ToList();
        }

        [HttpGet]
        public List<EMKB> GetEMKBList()
        {
            string sql = "exec sp_emkb";
            return connEM.Query<EMKB>(sql).ToList();
        }

        [HttpGet]
        public List<Config> GetConfigBySystemName(string SystemName)
        {
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[Boards] WHERE TesterID = '" + SystemName + "' ORDER BY Slot";
            return conn.Query<Config>(sql).ToList();
        }

        [HttpGet]
        public List<Board> GetBoardsDetail(string Sn)
        {
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[Failuar] WHERE SN = '" + Sn + "' ORDER BY Date";
            return conn.Query<Board>(sql).ToList();
        }

        [HttpGet]
        public List<PCBAShortage> GetShortageBySlot(string slot)
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE] WHERE IsReceived = 0 AND Slot = '" + slot + "'";
            return conn.Query<PCBAShortage>(sql).ToList();
        }

        [HttpGet]
        public List<Instrument> GetInstrumentsStatus()
        {
            string sql = "SELECT * FROM [PNINOUT].[dbo].[Instrument] ORDER BY PN";
            return conn.Query<Instrument>(sql).ToList();
        }

        [HttpGet]
        public List<FA> GetFAList()
        {
            string sql = "SELECT * FROM [192.168.163.1].[FCT_Test].[dbo].[IKW_FA_MaterialTracking]";
            return conn.Query<FA>(sql).ToList();
        }

        [HttpGet]
        public List<PCBAShortage> GetShortageList()
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE]";
            return conn.Query<PCBAShortage>(sql).ToList();
        }
    }
}
