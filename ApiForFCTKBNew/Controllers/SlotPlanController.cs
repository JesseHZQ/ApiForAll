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
    }
}
