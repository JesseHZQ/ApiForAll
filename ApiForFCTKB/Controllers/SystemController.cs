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
using System.IO;
using ApiForFCTKB.Models;

namespace ApiForFCTKB.Controllers
{
    public partial class SystemController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        /// <summary>
        /// 查询kanban
        /// </summary>
        [HttpGet]
        public Resp GetUFSystem(string type)
        {
            string sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ISNULL(ShippingDate, 0) = '0'";
            string sqlshortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE SLOT LIKE ('" + type + "%') AND IsReceived = 0";
            string sqlconfig = "SELECT * FROM KANBAN_SLOTCONFIG WHERE SLOT LIKE ('" + type + "%')";
            if (type == "IF")
            {
                sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE IN ( 'IF', 'MF') AND ISNULL(ShippingDate, 0) = '0'";
                sqlshortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE (SLOT LIKE ('IF%') OR SLOT LIKE('MF%')) AND IsReceived = 0";
                sqlconfig = "SELECT * FROM KANBAN_SLOTCONFIG WHERE (SLOT LIKE ('IF%') OR SLOT LIKE('MF%'))";
            }
            Resp resp = new Resp();
            resp.slotPlans = conn.Query<SlotPlan>(sqlplan).ToList();
            resp.slotShortages = conn.Query<SlotShortage>(sqlshortage).ToList();
            resp.slotConfigs = conn.Query<SlotConfig>(sqlconfig).ToList();
            return resp;
        }

        /// <summary>
        /// 刷新看板
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string RefreshSystem()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            List<SlotPlan> list_slotplan = GetSlotPlan();
            List<SlotPO> list_slotpo = GetPO();
            List<SlotCORCIssue> list_corcissue = GetCORCIssue();
            List<SlotShortage> list_shortage = GetShortage();
            List<SlotConfig> list_config = GetConfig(list_slotplan);

            // 遍历SlotPlan中所有的Slot
            foreach (SlotPlan slot in list_slotplan)
            {
                // 遍历po匹配
                foreach (SlotPO po in list_slotpo)
                {
                    if (slot.Slot == po.Slot)
                    {
                        slot.PO = po.PO;
                    }
                }

                // 遍历corc匹配
                foreach (SlotCORCIssue corc in list_corcissue)
                {
                    if (slot.Slot == corc.Slot)
                    {
                        slot.CORC = corc.CorcStatus;
                        slot.CORC_Issue = corc.CorcIssue;
                    }
                }

                // 添加默认的一个0-pin
                SlotConfig pin = new SlotConfig();
                pin.Slot = slot.Slot;
                pin.PN = "";
                pin.Description = "0-Pin";
                pin.QTY = "1";
                pin.DelayTips = "";
                pin.IsReady = true;

                list_config.Add(pin);

                Status status = new Status();
                if (slot.Slot.IndexOf("UF") > -1)
                {
                    status = GetUFStatus(slot.Slot);
                }
                slot.Launch = status.Launch;
                slot.CoreBU = status.CorBU;
                slot.PV = status.PV;
                slot.OI = status.OI;
                slot.TestBU = status.TestBU;
                slot.CSW = status.CSW;
                slot.QFAA = status.QFAA;
                slot.BU = status.BU;
                slot.Pack = status.Pack;

                string sql = "SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Boards] WHERE TESTERID = @TESTERID";
                MinioneConfig mc = new MinioneConfig();
                mc.TesterID = slot.Slot;
                List<MinioneConfig> MinioneList = connMinione.Query<MinioneConfig>(sql, mc).ToList();
                if (MinioneList.Count > 0)
                {
                    slot.IsLoad = true;
                }
                else
                {
                    slot.IsLoad = false;
                }
            }

            // 插入更新Slotplan
            InsertOrUpdateSlotPlan(list_slotplan);

            // 插入更新Shortage
            InsertOrUpdateSlotShortage(list_shortage);

            // 插入更新config
            InsertOrUpdateSlotConfig(list_config);

            return "OK";
        }

        /// <summary>
        /// 获取Slot Plan
        /// </summary>
        public List<SlotPlan> GetSlotPlan()
        {
            // 读取slot plan
            List<SlotPlan> list = new List<SlotPlan>();
            // slot plan 文件夹路径
            string dirPath = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\Slot Plan";
            // 定义slot type: UF IF MF D
            string slotType = "";
            DirectoryInfo dir = new DirectoryInfo(dirPath);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string fileParh = "";
                if (file.FullName.Contains("UFLEX"))
                {
                    fileParh = file.FullName;
                    slotType = "UF";
                }
                if (file.FullName.Contains("IFLEX"))
                {
                    fileParh = file.FullName;
                    slotType = "IF";
                }
                if (file.FullName.Contains("MFLEX"))
                {
                    fileParh = file.FullName;
                    slotType = "MF";
                }
                if (file.FullName.Contains("Dragon"))
                {
                    fileParh = file.FullName;
                    slotType = "D";
                }
                if (fileParh != "")
                {
                    Workbook wb = new Workbook(fileParh);
                    Cells cells = wb.Worksheets["Sheet1"].Cells;
                    DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                    foreach (DataRow item in dt.Rows)
                    {
                        int week = Tool.WeekOfYear(DateTime.Now); // 获取当前周别
                        float mrp = float.Parse(item[dt.Columns.IndexOf("MRP")].ToString()); // 获取MRP
                        int range = 5; // 抓取的周范围 后期改成可修改
                        bool wkRange = true; // 定义周范围的Bool
                        if (week + range <= 54) // 年底之前逻辑简单
                        {
                            wkRange = (mrp >= week && mrp <= week + range);
                        }
                        else // 年底的时候周别逻辑
                        {
                            wkRange = (mrp >= week && mrp <= 54) || (mrp >= 1 && mrp <= week + range - 53);
                        }
                        // TST == STS 并且 抓取5周内的数据
                        if (wkRange && item[dt.Columns.IndexOf("TST")].ToString() == "STS")
                        {
                            SlotPlan info = new SlotPlan();
                            info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                            info.Type = slotType;
                            info.Model = item[dt.Columns.IndexOf("Model")].ToString();
                            info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                            info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                            info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                            info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                            list.Add(info);
                        }
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// 插入或更新Slot Plan
        /// </summary>
        /// <param name="slotplan"></param>
        public void InsertOrUpdateSlotPlan(List<SlotPlan> slotplan)
        {
            foreach (SlotPlan item in slotplan)
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE SLOT = @SLOT";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query, item).ToList();
                if (slotlist.Count > 0)
                {
                    string update = "UPDATE KANBAN_SLOTPLAN SET TYPE = @TYPE, MODEL = @MODEL, CUSTOMER = @CUSTOMER, PO = @PO, SO = @SO, MRP = @MRP, PD = @PD, CORC = @CORC, Launch = @Launch, CoreBU = @CoreBU, PV = @PV, OI = @OI, TestBU = @TestBU, CSW = @CSW, QFAA = @QFAA, BU = @BU, Pack = @Pack, CORC_Issue = @CORC_Issue, IsLoad = @IsLoad WHERE SLOT = @SLOT";
                    conn.Execute(update, item);
                }
                else
                {
                    string update = "INSERT INTO KANBAN_SLOTPLAN (SLOT, TYPE, MODEL, CUSTOMER, PO, SO, MRP, PD, CORC, Launch, CoreBU, PV, OI, TestBU, CSW, QFAA, BU, Pack, CORC_Issue, IsLoad) VALUES (@SLOT, @TYPE, @MODEL, @CUSTOMER, @PO, @SO, @MRP, @PD, @CORC, @Launch, @CoreBU, @PV, @OI, @TestBU, @CSW, @QFAA, @BU, @Pack, @CORC_Issue, @IsLoad)";
                    conn.Execute(update, item);
                }
            }
        }

        /// <summary>
        /// 获取所有的PO
        /// </summary>
        public List<SlotPO> GetPO()
        {
            // 读取open_po
            string path = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\Open_PO";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("Open_PO"))
                {
                    path = file.FullName;
                    break;
                }
            }

            List<SlotPO> list = new List<SlotPO>();
            Workbook wb = new Workbook(path);
            foreach (Worksheet ws in wb.Worksheets)
            {
                Cells cells = ws.Cells;
                DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
                foreach (DataRow item in dt.Rows)
                {
                    string slot = item[dt.Columns.IndexOf("Slot Number\\Misc Order")].ToString();
                    string po = item[dt.Columns.IndexOf("PO#")].ToString();
                    bool isSysOrder = item[dt.Columns.IndexOf("Supplier note")].ToString().IndexOf("System Order") > -1;
                    if (slot != "" && isSysOrder)
                    {
                        if (dic.Keys.Contains(slot))
                        {
                            dic[slot].Add(po);
                        }
                        else
                        {
                            List<string> polist = new List<string>();
                            polist.Add(po);
                            dic.Add(slot, polist);
                        }
                    }
                }
                foreach (var item in dic)
                {
                    SlotPO slotPO = new SlotPO();
                    slotPO.Slot = item.Key;
                    item.Value.Distinct().ToList().ForEach(i =>
                    {
                        if (slotPO.PO == "" || slotPO.PO == null)
                        {
                            slotPO.PO = i;
                        }
                        else
                        {
                            slotPO.PO = slotPO.PO + "," + i;
                        }
                    });
                    list.Add(slotPO);
                }
            }
            return list;
        }

        /// <summary>
        /// 插入或更新Shortage
        /// </summary>
        /// <param name="slotshortage"></param>
        public void InsertOrUpdateSlotShortage(List<SlotShortage> slotshortage)
        {
            foreach (SlotShortage item in slotshortage)
            {
                string query = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE SLOT = @SLOT AND PN = @PN";
                List<SlotShortage> list = conn.Query<SlotShortage>(query, item).ToList();
                if (list.Count > 0)
                {
                    string update = "UPDATE KANBAN_SLOTSHORTAGE SET QTY = @QTY WHERE SLOT = @SLOT AND PN = @PN";
                    conn.Execute(update, item);
                }
                else
                {
                    string update = "INSERT INTO KANBAN_SLOTSHORTAGE VALUES (@SLOT, @PN, @Description, @QTY, @ETA, @IsReceived)";
                    conn.Execute(update, item);
                }
            }
        }

        /// <summary>
        /// 插入或更新config
        /// </summary>
        /// <param name="slotconfig"></param>
        public void InsertOrUpdateSlotConfig(List<SlotConfig> slotconfig)
        {
            foreach (SlotConfig item in slotconfig)
            {
                string query = "SELECT * FROM KANBAN_SLOTCONFIG WHERE SLOT = @SLOT AND PN = @PN";
                List<SlotShortage> list = conn.Query<SlotShortage>(query, item).ToList();
                if (list.Count > 0)
                {
                    string update = "UPDATE KANBAN_SLOTCONFIG SET QTY = @QTY WHERE SLOT = @SLOT AND PN = @PN";
                    conn.Execute(update, item);
                }
                else
                {
                    string update = "INSERT INTO KANBAN_SLOTCONFIG VALUES (@SLOT, @PN, @Description, @QTY, @DelayTips, @IsReady)";
                    conn.Execute(update, item);
                }
            }
        }

        /// <summary>
        /// 获取CORC&CORC ISSUE
        /// </summary>
        public List<SlotCORCIssue> GetCORCIssue()
        {
            string path = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\SO&CORC Issue Report";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("SO&CORC"))
                {
                    path = file.FullName;
                    break;
                }
            }
            List<SlotCORCIssue> list = new List<SlotCORCIssue>();
            Workbook wb = new Workbook(path);
            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name.IndexOf("Open issue") > -1)
                {
                    Cells cells = ws.Cells;
                    DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                    foreach (DataRow item in dt.Rows)
                    {
                        SlotCORCIssue info = new SlotCORCIssue();
                        info.Slot = item[dt.Columns.IndexOf("Slot")].ToString();
                        info.CorcStatus = item[dt.Columns.IndexOf("CORC status")].ToString();
                        info.CorcIssue = item[dt.Columns.IndexOf("SO or CORC open issue")].ToString();
                        list.Add(info);
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// 获取Shortage
        /// </summary>
        public List<SlotShortage> GetShortage()
        {
            string path = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\CSO";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("Material Shortage"))
                {
                    path = file.FullName;
                    break;
                }
            }
            List<SlotShortage> list = new List<SlotShortage>();
            Workbook wb = new Workbook(path);
            Cells cells = wb.Worksheets[0].Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                // Type = P  Shortage != 0  Slot not contain Misc
                if (item[dt.Columns.IndexOf("Type")].ToString() == "P" && item[dt.Columns.IndexOf("Shortage")].ToString() != "0" && item[dt.Columns.IndexOf("Slot")].ToString().IndexOf("Misc") == -1)
                {
                    SlotShortage info = new SlotShortage();
                    info.Slot = item[dt.Columns.IndexOf("Slot")].ToString();
                    info.PN = item[dt.Columns.IndexOf("Component")].ToString();
                    info.Description = item[dt.Columns.IndexOf("CompDescription")].ToString();
                    info.QTY = item[dt.Columns.IndexOf("ExtendedQty")].ToString();
                    info.ETA = item[dt.Columns.IndexOf("ConfirmedDate")].ToString();
                    info.IsReceived = false;
                    list.Add(info);
                }
            }
            return list;
        }

        /// <summary>
        /// 获取Shortage
        /// </summary>
        //public List<SlotConfig> GetConfig2()
        //{
        //    List<SlotConfig> list = new List<SlotConfig>();
        //    string path = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\E_KANBAN";
        //    DirectoryInfo dir = new DirectoryInfo(path);
        //    FileInfo[] files = dir.GetFiles();
        //    foreach (FileInfo file in files)
        //    {
        //        if (file.Name.IndexOf("kanban") > -1)
        //        {
        //            path = file.FullName;
        //            Workbook wb = new Workbook(path);
        //            Cells cells = wb.Worksheets[wb.Worksheets.Count - 1].Cells;
        //            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
        //            DataRow dr = dt.Rows[0];
        //            foreach (DataRow item in dt.Rows)
        //            {
        //                if (dt.Rows.IndexOf(item) != 0)
        //                {
        //                    foreach (DataColumn col in dt.Columns)
        //                    {
        //                        if (col.ColumnName.IndexOf("Column") == -1 && item[col.Ordinal].ToString() != "")
        //                        {
        //                            SlotConfig info = new SlotConfig();
        //                            info.Slot = item[0].ToString();
        //                            info.PN = col.ColumnName;
        //                            info.Description = dr[col.Ordinal].ToString();
        //                            info.QTY = item[col.Ordinal].ToString();
        //                            info.DelayTips = "";
        //                            info.IsReady = false;
        //                            list.Add(info);
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return list;
        //}

        public List<SlotConfig> GetConfig(List<SlotPlan> list_slotplan)
        {
            string[] PN_List = new string[] {
                "TDN-974-221-20",
                "TDN-974-221-30",
                "TDN-604-356-01",
                "TDN-604-356-03",
                "TDN-604-356-04",
                "TDN-805-052-05",
                "TDN-974-294-30",
                "TDN-604-375-02",
                "TDN-605-743-01",
                "TDN-974-217-30",
                "TDN-974-390-02",
                "TDN-974-230-00",
                "TDN-974-232-00",
                "TDN-630-036-30",
                "TDN-631-938-02",
                "TDN-632-627-01",
                "TDN-632-629-01",
                "TDN-633-246-00",
                "TDN-636-752-01",
                "TDN-639-210-01",
                "TDN-974-296-30",
                "TDN-632-859-21",
                "TDN-633-675-03",
                "TDN-630-035-40",
                "TDN-617-743-00",
                "TDN-617-744-00",
                "TDN-974-245-40",
                "TDN-604-375-12",
                "TDN-625-393-51",
                "TDN-639-209-02",
                "TDN-654-990-00",
                //"TDN-202-000-01",
                //"TDN-866-999-02",
                //"TDN-866-911-10",
                //"TDN-866-653-00",
                "TDN-805-002-60",
                "TDN-805-003-11",
                "TDN-805-003-50",
                "TDN-805-229-61",
                "TDN-805-230-60",
                "TDN-805-251-50",
                "TDN-805-333-50",
                "TDN-805-386-01",
                "TDN-805-740-50",
                "TDN-949-974-50"
            };
            List<SlotConfig> list = new List<SlotConfig>();
            string path = @"\\10.194.51.14\backup\jesse\FCT E-KANBAN Database\E-Slot";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("Slot"))
                {
                    path = file.FullName;
                    break;
                }
            }
            Workbook wb = new Workbook(path);
            Cells cells = wb.Worksheets["Slot"].Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                if (list_slotplan.FindAll(x => x.Slot == item["Slot"].ToString()).Count > 0 && PN_List.Contains(item["Component"].ToString()))
                {
                    SlotConfig info = new SlotConfig();
                    info.Slot = item["Slot"].ToString();
                    info.PN = item["Component"].ToString().Substring(4);
                    info.QTY = item["Extended Qty"].ToString();
                    info.DelayTips = "";
                    info.IsReady = false;
                    list.Add(info);
                }
            }
            return list;
        }

        /// <summary>
        /// 获取系统的单个状态的方法
        /// </summary>
        /// <param name="slot"></param>
        /// <param name="station"></param>
        /// <returns></returns>
        public string GetStatusByItem(string slot, string item)
        {
            string sql = "SELECT top 1 SUBSTRING(CONVERT(varchar, DATENAME(YEAR,Date)),3,2) + " +
                         "RIGHT('00' + CAST(DATENAME(WEEK, a.Date) AS nvarchar(50)), 2) + '.' + " +
                         "case when convert(varchar(50), datepart(dw, A.Date) - 1) = '0' then '7' else convert(varchar(50), datepart(dw, A.Date) - 1) end as WeekNum, b.systemName as Slot " +
                         "FROM [SC3_test].[dbo].[Check_SC3] a " +
                         "join [SC3_test].[dbo].[SystemProperties] b " +
                         "on a.systemId = b.systemId " +
                         "where b.systemName = '" + slot + "' and a.CheckItemDescription like '%" + item + "%' " +
                         "order by a.Date";
            List<SingleStatus> list = conn.Query<SingleStatus>(sql).ToList();
            if (list.Count >= 1)
            {
                return list[0].WeekNum;
            }
            return "";
        }

        /// <summary>
        /// 获取系统的单个状态的方法
        /// </summary>
        /// <param name="slot"></param>
        /// <param name="station"></param>
        /// <returns></returns>
        public string GetStatusByStation(string slot, string station)
        {
            string sql = "SELECT top 1 SUBSTRING(CONVERT(varchar, DATENAME(YEAR,Date)),3,2) + " +
                         "RIGHT('00' + CAST(DATENAME(WEEK, a.Date) AS nvarchar(50)), 2) + '.' + " +
                         "case when convert(varchar(50), datepart(dw, A.Date) - 1) = '0' then '7' else convert(varchar(50), datepart(dw, A.Date) - 1) end as WeekNum, b.systemName as Slot " +
                         "FROM [SC3_test].[dbo].[Check_SC3] a " +
                         "join [SC3_test].[dbo].[SystemProperties] b " +
                         "on a.systemId = b.systemId " +
                         "where b.systemName = '" + slot + "' and a.ListBy like '%" + station + "%' " +
                         "order by a.Date";
            List<SingleStatus> list = conn.Query<SingleStatus>(sql).ToList();
            if (list.Count >= 1)
            {
                return list[0].WeekNum;
            }
            return "";
        }

        /// <summary>
        /// 获取系统整个状态的对象
        /// </summary>
        /// <param name="slot"></param>
        /// <returns></returns>
        public Status GetUFStatus(string slot)
        {
            Status status = new Status();
            status.Slot = slot;
            status.Launch = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk");
            status.CorBU = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk");
            status.PV = GetStatusByItem(slot, "1. QDC Latch(inside test head)");
            status.OI = GetStatusByItem(slot, "33. Clamp(To heat exchange)");
            status.TestBU = GetStatusByItem(slot, @"1.1 Verify latest PO and CORC.Save CORC and Slot Locator in below path: \\ Agate\UltraFlex Project\Sync data\Shipped\CORC Tracking");
            status.CSW = GetStatusByItem(slot, "1.28 Verify limit table and firmware updated before CSW");
            status.QFAA = GetStatusByItem(slot, "1.1 Record HFE Level");
            status.BU = GetStatusByItem(slot, "1.2确认门两侧的接地标签处接地线固定完好.");
            status.Pack = GetStatusByItem(slot, "1. 木箱检查：检查没有多余文字或者图，没有大面积划伤（图1）；检查没有多余的毛刺毛边杂物，如果有需要清理干净（图2,3,4）。检查木箱所有钉子已经钉好，没有外漏，如果有需要拔出来重新钉（图5,6,7）。");
            return status;
        }
    }
}