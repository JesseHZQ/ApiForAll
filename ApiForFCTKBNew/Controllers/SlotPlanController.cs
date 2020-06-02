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
    public class SlotPlanController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public IDbConnection connPNINOUT = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        [HttpGet]
        public string UpdateSlotPlan()
        {
            try
            {
                List<SlotPlan> slotplan = GetSlotPlan();

                string query = "SELECT * FROM KANBAN_SLOTPLAN";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
                foreach (SlotPlan item in slotplan)
                {
                    SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    if (sp != null)
                    {
                        string update = "UPDATE KANBAN_SLOTPLAN SET TYPE = @TYPE, MODEL = @MODEL, CUSTOMER = @CUSTOMER, SO = @SO, MRP = @MRP, PD = @PD, LastUpdatedTime = @LastUpdatedTime WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                    else
                    {
                        string insert = "INSERT INTO KANBAN_SLOTPLAN (SLOT, TYPE, MODEL, CUSTOMER, SO, MRP, PD, LastUpdatedTime) VALUES (@SLOT, @TYPE, @MODEL, @CUSTOMER, @SO, @MRP, @PD, @LastUpdatedTime)";
                        conn.Execute(insert, item);
                    }
                }
                // todo: 池子里有 Excel里没有的 如何处理
                // 修复了push out 的问题
                List<SlotPlan> slotplanall = GetSlotPlanAll();
                foreach (SlotPlan item in slotplanall)
                {
                    SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    if (sp != null)
                    {
                        string update = "UPDATE KANBAN_SLOTPLAN SET TYPE = @TYPE, MODEL = @MODEL, CUSTOMER = @CUSTOMER, SO = @SO, MRP = @MRP, PD = @PD, LastUpdatedTime = @LastUpdatedTime WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                }
                return "SlotPlan OK!";
            }
            catch (Exception ex)
            {
                return "SlotPlan Failed! " + ex.Message;
            }
        }

        public List<SlotPlan> GetSlotPlan()
        {
            List<SlotPlan> list = new List<SlotPlan>();
            string dirPath = @"\\10.194.51.14\TER\FCT E-KANBAN Database\Slot Plan";
            DirectoryInfo dir = new DirectoryInfo(dirPath);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string filePath = "";
                string slotType = "";
                // 过滤临时文件
                if (file.FullName.Contains("~"))
                {
                    continue;
                }
                if (file.FullName.Contains("UFLEX"))
                {
                    filePath = file.FullName;
                    slotType = "UF";
                }
                if (file.FullName.Contains("IFLEX"))
                {
                    filePath = file.FullName;
                }
                if (file.FullName.Contains("Dragon"))
                {
                    filePath = file.FullName;
                    slotType = "D";
                }
                if (file.FullName.Contains("J750"))
                {
                    filePath = file.FullName;
                    slotType = "J750";
                }
                if (filePath != "")
                {
                    // IF MF 混在一起 所以只能单独考虑
                    if (file.FullName.Contains("IFLEX"))
                    {
                        Workbook wb = new Workbook(filePath);
                        Cells cells = wb.Worksheets["IFLEX"].Cells;
                        slotType = "IF";
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            bool wkRange = Tool.GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
                            if (wkRange && item[dt.Columns.IndexOf("TST")].ToString() == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }

                        wb = new Workbook(filePath);
                        cells = wb.Worksheets["MFLEX"].Cells;
                        slotType = "MF";
                        dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            bool wkRange = Tool.GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
                            if (wkRange && item[dt.Columns.IndexOf("TST")].ToString() == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }
                    }
                    else
                    {
                        Workbook wb = new Workbook(filePath);
                        Cells cells = wb.Worksheets[0].Cells;
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            bool wkRange = item[dt.Columns.IndexOf("MRP")].ToString() == "" ? false : Tool.GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
                            string tst = slotType == "J750" ? item[dt.Columns.IndexOf("Site")].ToString() : item[dt.Columns.IndexOf("TST")].ToString();
                            if (wkRange && tst == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = slotType == "J750" ? item[dt.Columns.IndexOf("Model")].ToString() : item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }
                    }
                }
            }
            return list;
        }

        public List<SlotPlan> GetSlotPlanAll()
        {
            List<SlotPlan> list = new List<SlotPlan>();
            string dirPath = @"\\10.194.51.14\TER\FCT E-KANBAN Database\Slot Plan";
            DirectoryInfo dir = new DirectoryInfo(dirPath);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string filePath = "";
                string sheetName = "";
                string slotType = "";
                // 过滤临时文件
                if (file.FullName.Contains("~"))
                {
                    continue;
                }
                if (file.FullName.Contains("UFLEX"))
                {
                    filePath = file.FullName;
                    sheetName = "Current Week";
                    slotType = "UF";
                }
                if (file.FullName.Contains("IFLEX"))
                {
                    filePath = file.FullName;
                }
                if (file.FullName.Contains("Dragon"))
                {
                    filePath = file.FullName;
                    //sheetName = "Dragon";
                    slotType = "D";
                }
                if (file.FullName.Contains("J750"))
                {
                    filePath = file.FullName;
                    sheetName = "Sheet1";
                    slotType = "J750";
                }
                if (filePath != "")
                {
                    // IF MF 混在一起 所以只能单独考虑
                    if (file.FullName.Contains("IFLEX"))
                    {
                        Workbook wb = new Workbook(filePath);
                        Cells cells = wb.Worksheets["IFLEX"].Cells;
                        slotType = "IF";
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            if (item[dt.Columns.IndexOf("TST")].ToString() == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }

                        wb = new Workbook(filePath);
                        cells = wb.Worksheets["MFLEX"].Cells;
                        slotType = "MF";
                        dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            if (item[dt.Columns.IndexOf("TST")].ToString() == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }
                    }
                    else
                    {
                        Workbook wb = new Workbook(filePath);
                        Cells cells = null;
                        if (sheetName == "")
                        {
                            cells = wb.Worksheets[0].Cells; // sheetName的不稳定性 所以用索引
                        }
                        else
                        {
                            cells = wb.Worksheets[sheetName].Cells;
                        }
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            string tst = slotType == "J750" ? item[dt.Columns.IndexOf("Site")].ToString() : item[dt.Columns.IndexOf("TST")].ToString();
                            if (tst == "STS")
                            {
                                SlotPlan info = new SlotPlan();
                                info.Slot = item[dt.Columns.IndexOf("Slot #")].ToString();
                                info.Type = slotType;
                                info.Model = slotType == "J750" ? item[dt.Columns.IndexOf("Model")].ToString() : item[dt.Columns.IndexOf("Product Model")].ToString();
                                info.Customer = item[dt.Columns.IndexOf("Customer")].ToString();
                                info.SO = item[dt.Columns.IndexOf("S/O #")].ToString();
                                info.MRP = item[dt.Columns.IndexOf("MRP")].ToString();
                                info.PD = item[dt.Columns.IndexOf("PD")].ToString();
                                info.LastUpdatedTime = DateTime.Now;
                                list.Add(info);
                            }
                        }
                    }
                }
            }
            return list;
        }
        
        [HttpGet]
        public List<SlotPlan> GetSystemByType(string type)
        {
            string sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ShippingDate is null ORDER BY MRP, PlanShipDate, Slot";
            if (type == "IF")
            {
                sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE IN ( 'IF', 'MF') AND ShippingDate is null ORDER BY MRP, PlanShipDate, Slot";
            }
            if (type == "J750")
            {
                sqlplan = "SELECT * FROM KANBAN_SLOTPLAN A LEFT JOIN KANBAN_SLOTPLAN_J750 B ON A.SLOT = B.SLOT WHERE A.TYPE = 'J750' AND A.ShippingDate is null ORDER BY A.MRP, A.PlanShipDate, A.Slot";
            }
            List<SlotPlan> list = conn.Query<SlotPlan>(sqlplan).ToList().Where(x => Tool.GetSlotPlanValidate(x.MRP) == true).ToList();
            string sqlConfig = "SELECT * FROM KANBAN_SLOTCONFIG";
            List<SlotConfig> listConfig = conn.Query<SlotConfig>(sqlConfig).ToList();
            string sqlShortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE IsReceived = 0";
            List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
            string sqlEIssue = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Status <> 'Close'";
            List<SlotEIssue> listEIssue = conn.Query<SlotEIssue>(sqlEIssue).ToList();

            List<PNQTY> listPNQTY = GetPNQTY();
            foreach (SlotPlan slotplan in list)
            {
                slotplan.ConfigList = new List<SlotConfig>();
                slotplan.ShortageList = new List<SlotShortage>();
                slotplan.EIssueList = new List<SlotEIssue>();

                // 获取自带的PNQTY 以便扣除
                List<string> strArr = slotplan.LoadInfo.Split(',').ToList();

                foreach (SlotConfig item in listConfig)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        int qty = item.QTY;
                        // 先扣loadinfo
                        foreach (string str in strArr)
                        {
                            if (item.PN == str)
                            {
                                qty--;
                                if (qty < 0)
                                {
                                    qty = 0;
                                }
                            }
                        }
                        if (qty == 0)
                        {
                            item.IsReady = 1;
                        }
                        else
                        {
                            // 再扣supermarket
                            foreach (PNQTY pq in listPNQTY)
                            {
                                if (pq.PN == item.PN)
                                {
                                    int diff = qty - pq.SupermarketQTY;
                                    if (diff <= 0)
                                    {
                                        pq.SupermarketQTY = -diff;
                                        item.IsReady = 1;
                                    }
                                    else
                                    {
                                        pq.SupermarketQTY = 0;
                                        int diff1 = pq.LeanQTY - diff;
                                        if (diff1 >= 0)
                                        {
                                            pq.LeanQTY = diff1;
                                            item.IsReady = 2;
                                        }
                                        else
                                        {
                                            pq.LeanQTY = 0;
                                            item.IsReady = 0;
                                        }
                                    }
                                }
                            }
                        }
                        slotplan.ConfigList.Add(item);
                    }
                }

                foreach (SlotShortage item in listShortage)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ShortageList.Add(item);
                    }
                }

                foreach (SlotEIssue item in listEIssue)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.EIssueList.Add(item);
                    }
                }
            }
            return list;
        }

        [HttpGet]
        public List<SlotPlan> GetSystemByType1(string type)
        {
            string sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ShippingDate is null ORDER BY ShippingTime, MRP, PlanShipDate, Slot";
            if (type == "IF")
            {
                sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE IN ( 'IF', 'MF') AND ShippingDate is null ORDER BY ShippingTime, MRP, PlanShipDate, Slot";
            }
            if (type == "J750")
            {
                sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN A LEFT JOIN KANBAN_SLOTPLAN_J750 B ON A.SLOT = B.SLOT WHERE A.TYPE = 'J750' AND A.ShippingDate is null ORDER BY A.MRP, A.PlanShipDate, A.Slot";
            }
            List<SlotPlan> list1 = conn.Query<SlotPlan>(sqlplan1).ToList().Where(x => float.Parse(x.MRP) > 40).ToList();
            List<SlotPlan> list2 = conn.Query<SlotPlan>(sqlplan1).ToList().Where(x => float.Parse(x.MRP) <= 40).ToList();
            List<SlotPlan> list = list1.Concat(list2).ToList();
            list = list.Where(x => Tool.GetSlotPlanValidate(x.MRP) == true).ToList();
            string sqlConfig = "SELECT * FROM KANBAN_SLOTCONFIG";
            List<SlotConfig> listConfig = conn.Query<SlotConfig>(sqlConfig).ToList();
            string sqlShortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE IsReceived = 0";
            List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
            string sqlEIssue = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Status <> 'Close'";
            List<SlotEIssue> listEIssue = conn.Query<SlotEIssue>(sqlEIssue).ToList();
            foreach (SlotPlan slotplan in list)
            {
                slotplan.ConfigList = new List<SlotConfig>();
                slotplan.ShortageList = new List<SlotShortage>();
                slotplan.EIssueList = new List<SlotEIssue>();
                foreach (SlotConfig item in listConfig)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ConfigList.Add(item);
                    }
                }
                foreach (SlotShortage item in listShortage)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ShortageList.Add(item);
                    }
                }
                foreach (SlotEIssue item in listEIssue)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.EIssueList.Add(item);
                    }
                }
            }
            return list;
        }

        [HttpGet]
        public List<SlotPlan> GetShippedSystemByType(string type)
        {
            string sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ShippingDate is not null ORDER BY PD, Slot";
            if (type == "IF")
            {
                sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE IN ( 'IF', 'MF') AND ShippingDate is not null ORDER BY PD, Slot";
            }
            if (type == "D")
            {
                sqlplan1 = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ShippingDate is not null ORDER BY ShippingDate";
            }
            List<SlotPlan> list = conn.Query<SlotPlan>(sqlplan1).ToList();
            string sqlConfig = "SELECT * FROM KANBAN_SLOTCONFIG";
            List<SlotConfig> listConfig = conn.Query<SlotConfig>(sqlConfig).ToList();
            string sqlShortage = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE IsReceived = 0";
            List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
            string sqlEIssue = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Status <> 'Close'";
            List<SlotEIssue> listEIssue = conn.Query<SlotEIssue>(sqlEIssue).ToList();
            foreach (SlotPlan slotplan in list)
            {
                slotplan.ConfigList = new List<SlotConfig>();
                slotplan.ShortageList = new List<SlotShortage>();
                slotplan.EIssueList = new List<SlotEIssue>();
                foreach (SlotConfig item in listConfig)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ConfigList.Add(item);
                    }
                }
                foreach (SlotShortage item in listShortage)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.ShortageList.Add(item);
                    }
                }
                foreach (SlotEIssue item in listEIssue)
                {
                    if (item.Slot == slotplan.Slot)
                    {
                        slotplan.EIssueList.Add(item);
                    }
                }
            }
            return list;
        }

        public List<PNQTY> GetPNQTY()
        {
            List<PNQTY> list = new List<PNQTY>();
            
            // Supermarket
            string sqlSupermarket = "select * from Supermarket_Rack";
            List<Rack> RackList = connPNINOUT.Query<Rack>(sqlSupermarket).ToList();

            foreach (Rack rack in RackList)
            {
                List<string> PNList = rack.PNList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> SNList = rack.SNList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> SlotList = rack.SlotList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> DateList = rack.DateList.Split(new char[] { '"', ',' }, StringSplitOptions.RemoveEmptyEntries).Where(x => x != " ").ToList();
                for (int i = 0; i < PNList.Count; i++)
                {
                    PNQTY pq = list.Where(x => x.PN == PNList[i]).FirstOrDefault();
                    if (pq != null)
                    {
                        pq.SupermarketQTY++;
                    }
                    else
                    {
                        PNQTY pnqty = new PNQTY();
                        pnqty.PN = PNList[i];
                        pnqty.SupermarketQTY = 1;
                        list.Add(pnqty);
                    }
                }
            }
            

            // Lean
            string sqlLean = "SELECT PN FROM [172.21.194.214].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            List<string> listLean = connMinione.Query<string>(sqlLean).ToList();

            foreach (string str in listLean)
            {
                PNQTY pq = list.Where(x => x.PN == str).FirstOrDefault();
                if (pq != null)
                {
                    pq.LeanQTY++;
                }
                else
                {
                    PNQTY pnqty = new PNQTY();
                    pnqty.PN = str;
                    pnqty.LeanQTY = 1;
                    list.Add(pnqty);
                }
            }

            // PSBSSB
            
            string sqlPSBSSB = "select * from KANBAN_PSBSSB";
            PSBSSB ps = conn.Query<PSBSSB>(sqlPSBSSB).FirstOrDefault();
            PNQTY psb = list.Where(x => x.PN == "974-221-20").FirstOrDefault();
            if(psb != null)
            {
                int diff = psb.LeanQTY - ps.PSBQty;
                psb.LeanQTY = diff <= 0 ? 0 : diff;
            }
            PNQTY ssb = list.Where(x => x.PN == "974-221-30").FirstOrDefault();
            if (ssb != null)
            {
                int diff = ssb.LeanQTY - ps.SSBQty;
                ssb.LeanQTY = diff <= 0 ? 0 : diff;
            }

            // 2020-05-21 特殊扣除56 UTAH 04
            PNQTY utah = list.Where(x => x.PN == "604-356-04").FirstOrDefault();
            utah.LeanQTY = (utah.LeanQTY - 68 >= 0 ? utah.LeanQTY - 68 : 0);

            return list;
        }



        // 多个增删改查
        [HttpPost]
        public int UpdateSlotPlan(SlotPlan slotplan)
        {
            string sql = "UPDATE KANBAN_SLOTPLAN SET PlanShipDate = @PlanShipDate, ShippingType = @ShippingType, ShippingDate = @ShippingDate, Launch = @Launch, Remark=@Remark, GroupNum = @GroupNum WHERE Slot = @Slot";
            return conn.Execute(sql, slotplan);
        }

        [HttpPost]
        public int UpdateShipping(SlotPlan slotplan)
        {
            slotplan.ShippingTime = DateTime.Now;
            string sql = "UPDATE KANBAN_SLOTPLAN SET PlanShipDate = @PlanShipDate, ShippingType = @ShippingType, ShippingDate = @ShippingDate, ShippingTime = @ShippingTime, Launch = @Launch, Remark=@Remark, GroupNum = @GroupNum WHERE Slot = @Slot";
            return conn.Execute(sql, slotplan);
        }

        [HttpPost]
        public int UpdateSlotPlanJ750(SlotPlan slotplan)
        {
            string sql = "UPDATE KANBAN_SLOTPLAN SET PlanShipDate = @PlanShipDate, ShippingType = @ShippingType, ShippingDate = @ShippingDate, Launch = @Launch, Remark=@Remark, GroupNum = @GroupNum WHERE Slot = @Slot";
            string s = "select * from KANBAN_SLOTPLAN_J750 where Slot = @Slot";
            List<SlotPlan> list = conn.Query<SlotPlan>(s, slotplan).ToList();
            string sqlJ750 = "";
            if (list.Count > 0)
            {
                sqlJ750 = @"UPDATE KANBAN_SLOTPLAN_J750
                     SET [WO] = @WO
                        ,[J_Launch] = @J_Launch
                        ,[J_Assy] = @J_Assy
                        ,[J_SST] = @J_SST
                        ,[J_OutGoing] = @J_OutGoing
                        ,[J_KanBan] = @J_KanBan
                        ,[J_BringUp] = @J_BringUp
                        ,[J_Heat] = @J_Heat
                        ,[J_Cold] = @J_Cold
                        ,[J_CSW] = @J_CSW
                        ,[J_QFAA] = @J_QFAA
                        ,[J_ButtonUp] = @J_ButtonUp
                        ,[J_Packing] = @J_Packing
                    WHERE [Slot] = @Slot";
            }
            else
            {
                sqlJ750 = @"INSERT INTO KANBAN_SLOTPLAN_J750
                    ([Slot]
                    ,[WO]
                    ,[J_Launch]
                    ,[J_Assy]
                    ,[J_SST]
                    ,[J_OutGoing]
                    ,[J_KanBan]
                    ,[J_BringUp]
                    ,[J_Heat]
                    ,[J_Cold]
                    ,[J_CSW]
                    ,[J_QFAA]
                    ,[J_ButtonUp]
                    ,[J_Packing]) VALUES
                    (@Slot
                    ,@WO
                    ,@J_Launch
                    ,@J_Assy
                    ,@J_SST
                    ,@J_OutGoing
                    ,@J_KanBan
                    ,@J_BringUp
                    ,@J_Heat
                    ,@J_Cold
                    ,@J_CSW
                    ,@J_QFAA
                    ,@J_ButtonUp
                    ,@J_Packing)";
            }
            return conn.Execute(sql, slotplan) + conn.Execute(sqlJ750, slotplan);
        }

        [HttpPost]
        public int DeleteSlotPlan(SlotPlan slotplan)
        {
            string sql = "DELETE KANBAN_SLOTPLAN WHERE ID = @ID";
            return conn.Execute(sql, slotplan);
        }

        [HttpPost]
        public string UpdateEIssue(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET Engineering_Issue = @Engineering_Issue WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdateRemark(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET Remark = @Remark WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdatePlanShip(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET PlanShipDate = @PlanShipDate WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdateShippingDate(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET ShippingDate = @ShippingDate WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdateGroup(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET GroupNum = @GroupNum WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdateLaunch(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET Launch = @Launch WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }


        public class Board
        {
            public string PN { get; set; }
            public string SN { get; set; }
            public string RackName { get; set; }
            public string Slot { get; set; }
            public DateTime Date { get; set; }
            public bool IsRefurbish { get; set; }
        }

        public class Rack
        {
            public int RackId { get; set; }
            public string RackName { get; set; }
            public int RackSize { get; set; }
            public int TypeId { get; set; }
            public string PN { get; set; }
            public string PNList { get; set; }
            public string SNList { get; set; }
            public string SlotList { get; set; }
            public string DateList { get; set; }
        }

        public class PNQTY
        {
            public string PN { get; set; }
            public int SupermarketQTY { get; set; }
            public int LeanQTY { get; set; }
        }

        private class PSBSSB
        {
            public int PSBQty { get; set; }
            public int SSBQty { get; set; }
        }

    }
}