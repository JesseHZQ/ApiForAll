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
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);
        public IDbConnection connSlot = new SqlConnection(ConfigurationManager.ConnectionStrings["SlotKB"].ConnectionString);

        [HttpGet]
        public List<LED> GetLEDList()
        {
            string sqlLED = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[LED]";
            List<OLDLED> list = conn.Query<OLDLED>(sqlLED).ToList();
            string sqlUpdate = "Update CT_LEDMaster SET SystemName = @System, Station = @Station, ShippingTime = @Ship WHERE BayName = @BayNO";
            int result = conn.Execute(sqlUpdate, list);
            string sql = "SELECT * FROM CT_LEDMaster A " +
                         "LEFT JOIN [SlotKB].[dbo].[KANBAN_SLOTPLAN] B " +
                         "ON A.SystemName = B.Slot " +
                         "LEFT JOIN CT_User C " +
                         "ON A.OwnerId = C.UserId " +
                         "ORDER BY A.Item";
            List<LED> l = conn.Query<LED>(sql).ToList();
            string sqlEIssue = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Status <> 'Close'";
            List<EIssue> listEIssue = connSlot.Query<EIssue>(sqlEIssue).ToList();
            foreach (LED led in l)
            {
                led.EIssueList = new List<EIssue>();
                
                foreach (EIssue item in listEIssue)
                {
                    if (item.Slot == led.Slot)
                    {
                        led.EIssueList.Add(item);
                    }
                }
            }
            return l;
        }

        [HttpGet]
        public List<LEDDetail> GetLEDDetailBySlot(string slot)
        {
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[BigData] WHERE SN = '" + slot + "' ORDER BY BeginTime";
            return conn.Query<LEDDetail>(sql).ToList();
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
            string sql = "SELECT a.* FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE] a join [SlotKB].[dbo].[KANBAN_SLOTPLAN] b on a.Slot = b.Slot where a.IsReceived = 0 and b.ShippingDate is null";
            return conn.Query<PCBAShortage>(sql).ToList();
        }

        [HttpGet]
        public List<EIssue> GetEIssueList()
        {
            string sql = "SELECT a.* FROM [SlotKB].[dbo].[KANBAN_SLOTEISSUE] a join [SlotKB].[dbo].[KANBAN_SLOTPLAN] b on a.Slot = b.Slot where a.Status = 'Open' and b.ShippingDate is null";
            return conn.Query<EIssue>(sql).ToList();
        }

        [HttpGet]
        public TempHumid GetRealTimeTHData()
        {
            string sql = "SELECT TOP 1 * FROM [FF_EDW].[dbo].[EDW_TEMPERATURE_HUMIDITY] where Location = 'FCT' ORDER BY CreationTime DESC";
            return connMinione.Query<TempHumid>(sql).SingleOrDefault();
        }
    }
}
