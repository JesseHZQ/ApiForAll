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
            //string sqlLED = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[LED]";
            //List<OLDLED> list = conn.Query<OLDLED>(sqlLED).ToList();
            //string sqlUpdate = "Update CT_LEDMaster SET SystemName = @System, Station = @Station, ShippingTime = @Ship WHERE BayName = @BayNO";
            //int result = conn.Execute(sqlUpdate, list);
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
            string sql = "SELECT * FROM [ControlTower].[dbo].[CT_LEDHistory] WHERE SystemName = '" + slot + "' ORDER BY FinishTime DESC";
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
            string sql = "SELECT t2.EngAssyStep,t1.EndTime FROM Steplog t1,AssemblyStep t2 WHERE Aslog = " + LogID + " and EndTime is not null AND t1.StepID=t2.ID";
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
            string SystemLong = SystemName;
            if (SystemName.Contains('F') && SystemName.StartsWith("1"))
            {
                SystemName = SystemName.Substring(1, 3) + SystemName.Substring(5, 3);
            }
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[Boards] WHERE TesterID in ('" + SystemName + "', '" + SystemLong + "') AND Slot is not null ORDER BY Slot";
            return conn.Query<Config>(sql).ToList();
        }

        [HttpGet]
        public List<FailBoard> GetFailureBySystemName(string SystemName)
        {
            string SystemLong = SystemName;
            if (SystemName.Contains('F') && SystemName.StartsWith("1"))
            {
                SystemName = SystemName.Substring(1, 3) + SystemName.Substring(5, 3);
            }
            string sql = "SELECT * FROM [192.168.163.1].[PCBA].[dbo].[Failuar] WHERE TesterID in ('" + SystemName + "', '" + SystemLong + "') AND NextLocation = 'Verifying'";
            return conn.Query<FailBoard>(sql).ToList();
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
            string sql = "SELECT * FROM [ControlTower].[dbo].[FA_Process_Part] where (FASN <> 'PCBA' or FASN is null) order by FASN desc";
            return conn.Query<FA>(sql).ToList();
        }

        [HttpGet]
        public List<PCBAShortage> GetShortageList()
        {
            string sql = "SELECT a.*, b.MRP FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE] a join [SlotKB].[dbo].[KANBAN_SLOTPLAN] b on a.Slot = b.Slot where a.IsReceived = 0 and b.ShippingDate is null and a.SupplierName is not null and a.SupplierName <> ''";
            return conn.Query<PCBAShortage>(sql).ToList().Where(x => GetValidate(x.MRP)).ToList();
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

        [HttpGet]
        public List<TempHumid> GetRealTimeTHDatas()
        {
            string sql = "SELECT TOP 1000 * FROM [FF_EDW].[dbo].[EDW_TEMPERATURE_HUMIDITY] where Location = 'FCT' ORDER BY CreationTime DESC";
            return connMinione.Query<TempHumid>(sql).ToList();
        }

        public bool GetValidate(string mrpStr)
        {
            int week = WeekOfYear(DateTime.Now); // 获取当周周别
            float mrp = float.Parse(mrpStr); // 获取MRP
            int range = 2; // 抓取的周范围 后期改成可修改
            bool IsRange = true; // 定义周范围的Bool
            if (week + range <= 54) // 年底之前逻辑简单
            {
                IsRange = (mrp >= week && mrp <= week + range);
            }
            else // 年底的时候周别逻辑
            {
                IsRange = (mrp >= week && mrp <= 54) || (mrp >= 1 && mrp <= week + range - 53);
            }
            return IsRange;
        }

        public static int WeekOfYear(DateTime dtime)
        {
            try
            {
                //确定此时间在一年中的位置
                int dayOfYear = dtime.DayOfYear;
                //当年第一天
                DateTime tempDate = new DateTime(dtime.Year, 1, 1);
                //确定当年第一天
                int tempDayOfWeek = (int)tempDate.DayOfWeek;
                tempDayOfWeek = tempDayOfWeek == 0 ? 7 : tempDayOfWeek;
                ////确定星期几
                int index = (int)dtime.DayOfWeek;
                index = index == 0 ? 7 : index;
                //当前周的范围
                DateTime retStartDay = dtime.AddDays(-(index - 1));
                DateTime retEndDay = dtime.AddDays(6 - index);
                //确定当前是第几周
                int weekIndex = (int)Math.Ceiling(((double)dayOfYear + tempDayOfWeek - 1) / 7);
                if (retStartDay.Year < retEndDay.Year)
                {
                    weekIndex = 1;
                }
                return weekIndex;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
