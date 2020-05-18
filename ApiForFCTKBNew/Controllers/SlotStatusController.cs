using ApiForFCTKBNew.Models;
using Aspose.Cells;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;


namespace ApiForFCTKBNew.Controllers
{
    public class SlotStatusController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public string UpdateSlotStatus()
        {
            try
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList().Where(x => x.Type == "UF").ToList();
                // Slot集合查询数据
                string strSlotList = "";
                foreach (SlotPlan slot in slotplan)
                {
                    strSlotList = strSlotList + "'" + slot.Slot + "',";
                }
                strSlotList = strSlotList.Substring(0, strSlotList.Length - 1);
                string sql24slot = "select a.SystemId, b.systemName, a.Date, a.ListBy, a.CheckItemDescription from [SC3_test].[dbo].[Check_SC3] a join [SC3_test].[dbo].[SystemProperties] b on a.SystemId = b.systemId where b.systemName in (" + strSlotList + ")";
                string sql12slot = "select a.SystemId, b.systemName, a.Date, a.ListBy, a.CheckItemDescription from [SC3_test].[dbo].[Check_12Slot] a join [SC3_test].[dbo].[SystemProperties] b on a.SystemId = b.systemId where b.systemName in (" + strSlotList + ")";

                List<CheckListItem> list24 = conn.Query<CheckListItem>(sql24slot).ToList();
                List<CheckListItem> list12 = conn.Query<CheckListItem>(sql12slot).ToList();

                foreach (SlotPlan item in slotplan)
                {
                    // 不能影响除12 24 之外的Model 
                    SlotStatus status = new SlotStatus();
                    if (item.Model.IndexOf("24") > -1)
                    {
                        status = GetUF24Status(item.Slot, list24);
                        item.CoreBU = status.CorBU;
                        item.PV = status.PV;
                        item.OI = status.OI;
                        item.TestBU = status.TestBU;
                        item.CSW = status.CSW;
                        item.QFAA = status.QFAA;
                        item.BU = status.BU;
                        item.Pack = status.Pack;
                    }
                    if (item.Model.IndexOf("12") > -1)
                    {
                        status = GetUF12Status(item.Slot, list12);
                        item.CoreBU = status.CorBU;
                        item.PV = status.PV;
                        item.OI = status.OI;
                        item.TestBU = status.TestBU;
                        item.CSW = status.CSW;
                        item.QFAA = status.QFAA;
                        item.BU = status.BU;
                        item.Pack = status.Pack;
                    }
                }
                string update = "UPDATE KANBAN_SLOTPLAN SET CoreBU = @CoreBU, PV = @PV, OI = @OI, TestBU = @TestBU, CSW = @CSW, QFAA = @QFAA, BU = @BU, Pack = @Pack WHERE ID = @ID";
                conn.Execute(update, slotplan);
                return "SlotStatus OK!";
            }
            catch (Exception ex)
            {
                return "SlotStatus Failed! " + ex.Message;
            }
        }

        public SlotStatus GetUF24Status(string slot, List<CheckListItem> list)
        {
            SlotStatus status = new SlotStatus();
            status.Slot = slot;
            status.Launch = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex SC3 Core Bring Up Station Checklist", list);
            status.CorBU = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex SC3 Core Bring Up Station Checklist", list);
            status.PV = GetStatusByItem(slot, "1. QDC Latch(inside test head)", "Ultraflex SC3 Core Bring Up Station Checklist", list);
            status.OI = GetStatusByItem(slot, "33. Clamp(To heat exchange)", "Ultraflex SC3 Core Bring Up Station Checklist", list);
            status.TestBU = GetStatusByItem(slot, @"1.1 Verify latest PO and CORC.Save CORC and Slot Locator in below path: \\ Agate\UltraFlex Project\Sync data\Shipped\CORC Tracking", "Ultraflex SC3 Test Bring Up Station Checklist", list);
            status.CSW = GetStatusByItem(slot, "Verify limit table and firmware updated before CSW", "Ultraflex SC3 Test Bring Up Station Checklist", list);
            status.QFAA = GetStatusByItem(slot, "1.1 Record HFE Level", "Ultraflex SC3 Final Station Checklist", list);
            status.BU = GetStatusByItem(slot, "1.1打开Test Head 的门,确认门两侧各有2个Danger 标签且无缺损脏污.", "Ultraflex SC3 Assembly Button up Checklist", list);
            status.Pack = GetStatusByItem(slot, "1. 木箱检查：检查没有多余文字或者图，没有大面积划伤（图1）；检查没有多余的毛刺毛边杂物，如果有需要清理干净（图2,3,4）。检查木箱所有钉子已经钉好，没有外漏，如果有需要拔出来重新钉（图5,6,7）。", "Ultraflex SC3 Packing Checklist", list);
            return status;
        }

        public SlotStatus GetUF12Status(string slot, List<CheckListItem> list)
        {
            SlotStatus status = new SlotStatus();
            status.Slot = slot;
            status.Launch = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", list);
            status.CorBU = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", list);
            status.PV = GetStatusByItem(slot, "1. QDC Latch(inside test head)", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", list);
            status.OI = GetStatusByItem(slot, "25. TH pipe drain QD(left)", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", list);
            status.TestBU = GetStatusByItem(slot, @"1.2 Verify latest PO and CorC. Save CORC and Slot Locator in below path: \\ Agate\UltraFlex Project\Sync data\Shipped\CORC Tracking", "Ultraflex 12Slot Test Bring Up Station Checklist", list);
            status.CSW = GetStatusByItem(slot, "1.29 Verify limit table and firmware updated before CSW", "Ultraflex 12Slot Test Bring Up Station Checklist", list);
            status.QFAA = GetStatusByItem(slot, "1.1 Record HFE Level", "UltraFlex 12Slot Final Station Checklist", list);
            status.BU = GetStatusByItem(slot, "1.1打开Test Head 的门,确认门侧有2个Dangerous 标签且无缺损脏污", "Ultraflex 12Slot Assembly Button up Checklist", list);
            status.Pack = GetStatusByItem(slot, "1. 木箱检查：检查没有多余文字或者图，没有大面积划伤（图1）；检查没有多余的毛刺毛边杂物，如果有需要清理干净（图2,3,4）。检查木箱所有钉子已经钉好，没有外漏，如果有需要拔出来重新钉（图5,6,7）。", "Ultraflex 12Slot Packing Checklist", list);
            return status;
        }

        public string GetStatusByItem(string slot, string item, string station, List<CheckListItem> list)
        {
            // x => x.SystemName.Contains(slot) 存在空格  不完全一样
            List<CheckListItem> checkListItemList = list.Where(x => x.SystemName.Contains(slot) && x.CheckItemDescription.Contains(item) && x.ListBy == station).ToList();
            CheckListItem checkListItem = new CheckListItem();
            if (checkListItemList.Count == 0)
            {
                checkListItem = null;
            }
            else
            {
                checkListItem = checkListItemList[0];
            }
            if (checkListItem != null && checkListItem.Date.Year > 2015)
            {
                int day = -1;
                switch (checkListItem.Date.DayOfWeek)
                {
                    case DayOfWeek.Sunday:
                        day = 7;
                        break;
                    case DayOfWeek.Monday:
                        day = 1;
                        break;
                    case DayOfWeek.Tuesday:
                        day = 2;
                        break;
                    case DayOfWeek.Wednesday:
                        day = 3;
                        break;
                    case DayOfWeek.Thursday:
                        day = 4;
                        break;
                    case DayOfWeek.Friday:
                        day = 5;
                        break;
                    case DayOfWeek.Saturday:
                        day = 6;
                        break;
                    default:
                        day = -1;
                        break;
                }
                string wk = Tool.WeekOfYear(checkListItem.Date).ToString();
                return checkListItem.Date.Year.ToString().Substring(2, 2) + (wk.Length > 1 ? wk : ("0" + wk)) + "." + day.ToString();
            }
            else
            {
                return "";
            }
        }

        public class CheckListItem
        {
            public DateTime Date { get; set; }
            public string ListBy { get; set; }
            public string CheckItemDescription { get; set; }
            public string SystemId { get; set; }
            public string SystemName { get; set; }
        }

        public class SlotStatus
        {
            public string Slot { get; set; }
            public string Launch { get; set; }
            public string CorBU { get; set; }
            public string PV { get; set; }
            public string OI { get; set; }
            public string TestBU { get; set; }
            public string CSW { get; set; }
            public string QFAA { get; set; }
            public string BU { get; set; }
            public string Pack { get; set; }
        }
    }
}
