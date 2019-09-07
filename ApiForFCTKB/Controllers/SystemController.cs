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
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace ApiForFCTKB.Controllers
{
    public partial class SystemController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        [HttpGet]
        public string Refresh()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");

            // 不超过一分钟
            InsertOrUpdateSlotPlan();
            UpdateSlotPO();
            UpdateSlotCORC();
            InsertOrUpdateSlotShortage();
            InsertOrUpdateSlotConfig();
            UpdateSlotMinioneStatus();
            UpdatePSBSSBQty();
            UpdateSlotStatus();

            return "OK";
        }

        [HttpGet]
        public PSBSSB GetPSBSSBQty()
        {
            string sql = "SELECT * FROM KANBAN_PSBSSB";
            return conn.Query<PSBSSB>(sql).SingleOrDefault();
        }

        [HttpGet]
        public List<SlotPlan> GetAllSystem()
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTPLAN] WHERE ISNULL(ShippingDate, 0) = '0' ORDER BY PD";
            return conn.Query<SlotPlan>(sql).ToList().Where(x => GetValidate(x.MRP) == true).ToList();
        }

        /// <summary>
        /// 生成Shift Report
        /// </summary>
        [HttpGet]
        public string generateExcel(string type)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\UFlex System Status.xls");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 2;
            string sql = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTPLAN] WHERE TYPE = '" + type + "' AND ShippingDate is null ORDER BY PD";
            List<SlotPlan> list = conn.Query<SlotPlan>(sql).ToList();
            list = list.Where(x => GetValidate(x.MRP) == true).ToList();
            foreach (SlotPlan s in list)
            {
                for (int i = 0; i < 18; i++)
                {
                    Style sty = cells[y, i].GetStyle();
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cells[y, i].SetStyle(sty);
                }
                cells[y, 0].PutValue(s.PD);
                cells[y, 1].PutValue(s.PlanShipDate);
                cells[y, 2].PutValue(s.Slot);
                cells[y, 3].PutValue(s.Model);
                cells[y, 4].PutValue(s.Customer);
                cells[y, 5].PutValue(s.SO);
                cells[y, 6].PutValue(s.PO);
                cells[y, 14].PutValue(s.Launch);
                cells[y, 15].PutValue(s.CORC);
                string sqlShortage = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE] WHERE SLOT = '" + s.Slot + "' AND IsReceived = 0";
                List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
                string Shortage_Issue = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage_Issue = Shortage_Issue + e.PN + "---" + e.QTY + ";";
                    }
                }
                cells[y, 16].PutValue(Shortage_Issue);
                string Engineering_Issue = "";
                if (s.Engineering_Issue != null && s.Engineering_Issue != "[]" && s.Engineering_Issue != "")
                {
                    var arrdata = Newtonsoft.Json.Linq.JArray.Parse(s.Engineering_Issue);
                    List<E_Issue> arr = arrdata.ToObject<List<E_Issue>>();
                    foreach (E_Issue e in arr)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item + ";";
                    }
                }
                cells[y, 17].PutValue(Engineering_Issue);
                Style style = new CellsFactory().CreateStyle();
                if (cells[y, 16].Value != null || cells[y, 17].Value != null)
                {
                    style.BackgroundColor = System.Drawing.Color.Red;
                    style.Pattern = BackgroundType.Solid;
                    style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                }
                else
                {
                    style.BackgroundColor = System.Drawing.Color.Green;
                    style.Pattern = BackgroundType.Solid;
                    style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                }
                if (s.Pack != "" && s.Pack != null)
                {
                    cells[y, 7].PutValue(s.Pack);
                    cells[y, 7].SetStyle(style);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cells[y, 8].PutValue(s.BU);
                    cells[y, 8].SetStyle(style);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cells[y, 9].PutValue(s.QFAA);
                    cells[y, 9].SetStyle(style);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cells[y, 10].PutValue(s.CSW);
                    cells[y, 10].SetStyle(style);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cells[y, 11].PutValue(s.TestBU);
                    cells[y, 11].SetStyle(style);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cells[y, 12].PutValue(s.OI);
                    cells[y, 12].SetStyle(style);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cells[y, 13].PutValue(s.CoreBU);
                    cells[y, 13].SetStyle(style);
                }
                y++;
            }

            Cells cellsShip = wb.Worksheets[1].Cells;
            int yShip = 2;
            string sqlShip = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTPLAN] WHERE TYPE = '" + type + "' AND ShippingDate is not null ORDER BY PD";
            List<SlotPlan> listShip = conn.Query<SlotPlan>(sqlShip).ToList();
            foreach (SlotPlan s in listShip)
            {
                for (int i = 0; i < 18; i++)
                {
                    Style sty = cellsShip[yShip, i].GetStyle();
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellsShip[yShip, i].SetStyle(sty);
                }
                cellsShip[yShip, 0].PutValue(s.PD);
                cellsShip[yShip, 1].PutValue(s.PlanShipDate);
                cellsShip[yShip, 2].PutValue(s.Slot);
                cellsShip[yShip, 3].PutValue(s.Model);
                cellsShip[yShip, 4].PutValue(s.Customer);
                cellsShip[yShip, 5].PutValue(s.SO);
                cellsShip[yShip, 6].PutValue(s.PO);
                cellsShip[yShip, 14].PutValue(s.Launch);
                cellsShip[yShip, 15].PutValue(s.CORC);
                string sqlShortage = "SELECT * FROM [SlotKB].[dbo].[KANBAN_SLOTSHORTAGE] WHERE SLOT = '" + s.Slot + "' AND IsReceived = 0";
                List<SlotShortage> listShortage = conn.Query<SlotShortage>(sqlShortage).ToList();
                string Shortage_Issue = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage_Issue = Shortage_Issue + e.PN + "---" + e.QTY + ";";
                    }
                }
                cellsShip[yShip, 16].PutValue(Shortage_Issue);
                string Engineering_Issue = "";
                if (s.Engineering_Issue != null && s.Engineering_Issue != "[]" && s.Engineering_Issue != "")
                {
                    var arrdata = Newtonsoft.Json.Linq.JArray.Parse(s.Engineering_Issue);
                    List<E_Issue> arr = arrdata.ToObject<List<E_Issue>>();
                    foreach (E_Issue e in arr)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item + ";";
                    }
                }
                cellsShip[yShip, 17].PutValue(Engineering_Issue);
                Style style = new CellsFactory().CreateStyle();
                if (cellsShip[yShip, 16].Value != null || cellsShip[yShip, 17].Value != null)
                {
                    style.BackgroundColor = System.Drawing.Color.Red;
                    style.Pattern = BackgroundType.Solid;
                    style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                }
                else
                {
                    style.BackgroundColor = System.Drawing.Color.Green;
                    style.Pattern = BackgroundType.Solid;
                    style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                }
                if (s.Pack != "" && s.Pack != null)
                {
                    cellsShip[yShip, 7].PutValue(s.Pack);
                    cellsShip[yShip, 7].SetStyle(style);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cellsShip[yShip, 8].PutValue(s.BU);
                    cellsShip[yShip, 8].SetStyle(style);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cellsShip[yShip, 9].PutValue(s.QFAA);
                    cellsShip[yShip, 9].SetStyle(style);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cellsShip[yShip, 10].PutValue(s.CSW);
                    cellsShip[yShip, 10].SetStyle(style);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cellsShip[yShip, 11].PutValue(s.TestBU);
                    cellsShip[yShip, 11].SetStyle(style);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cellsShip[yShip, 12].PutValue(s.OI);
                    cellsShip[yShip, 12].SetStyle(style);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cellsShip[yShip, 13].PutValue(s.CoreBU);
                    cellsShip[yShip, 13].SetStyle(style);
                }
                yShip++;
            }
            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\SR_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond + ".xlsx";
            wb.Save(filePath);
            return filePath;
        }

        /// <summary>
        /// 从服务器下载文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        [HttpGet]
        public HttpResponseMessage downloadFile(string filePath, string fileName)
        {
            try
            {
                var stream = new FileStream(filePath, FileMode.Open);
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = fileName
                };
                return response;
            }
            catch
            {
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
        }

        /// <summary>
        /// 查询kanban
        /// </summary>
        [HttpGet]
        public List<SlotPlan> GetSystemByType(string type)
        {
            string sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE = '" + type + "' AND ShippingDate is null ORDER BY PD, PlanShipDate, Slot";
            if (type == "IF")
            {
                sqlplan = "SELECT * FROM KANBAN_SLOTPLAN WHERE TYPE IN ( 'IF', 'MF') AND ShippingDate is null ORDER BY PD, PlanShipDate, Slot";
            }
            List<SlotPlan> list = conn.Query<SlotPlan>(sqlplan).ToList();
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
        



        /// <summary>
        /// 获取Slot Plan
        /// </summary>
        public List<SlotPlan> GetSlotPlan()
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
                    sheetName = "Dragon";
                    slotType = "D";
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
                            bool wkRange = GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
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
                            bool wkRange = GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
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
                        Cells cells = wb.Worksheets[sheetName].Cells;
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            bool wkRange = GetValidate(item[dt.Columns.IndexOf("MRP")].ToString());
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
                }
            }
            return list;
        }

        /// <summary>
        /// 插入或更新Slot Plan
        /// </summary>
        /// <param name="slotplan"></param>
        public void InsertOrUpdateSlotPlan()
        {
            List<SlotPlan> slotplan = GetSlotPlan();
            foreach (SlotPlan item in slotplan)
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
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
        }

        /// <summary>
        /// 获取所有的PO
        /// </summary>
        public List<SlotPO> GetSlotPO()
        {
            // 读取open_po
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\Open_PO";
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
        /// 更新SlotPlan中的PO
        /// </summary>
        public void UpdateSlotPO()
        {
            List<SlotPO> slotpo = GetSlotPO();
            foreach (SlotPO item in slotpo)
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
                SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                if (sp != null)
                {
                    string update = "UPDATE KANBAN_SLOTPLAN SET PO = @PO WHERE Slot = @Slot";
                    conn.Execute(update, item);
                }
            }
        }

        /// <summary>
        /// 获取CORC&CORC ISSUE
        /// </summary>
        public List<SlotCORC> GetSlotCORC()
        {
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\SO&CORC Issue Report";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("~"))
                {
                    continue;
                }
                if (file.FullName.Contains("SO&CORC"))
                {
                    path = file.FullName;
                    break;
                }
            }
            List<SlotCORC> list = new List<SlotCORC>();
            Workbook wb = new Workbook(path);
            foreach (Worksheet ws in wb.Worksheets)
            {
                if (ws.Name.IndexOf("Open issue") > -1)
                {
                    Cells cells = ws.Cells;
                    DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                    foreach (DataRow item in dt.Rows)
                    {
                        SlotCORC info = new SlotCORC();
                        info.Slot = item[dt.Columns.IndexOf("Slot")].ToString();
                        info.CORC = item[dt.Columns.IndexOf("CORC status")].ToString();
                        info.CORC_Issue = item[dt.Columns.IndexOf("SO or CORC open issue")].ToString();
                        list.Add(info);
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// 更新SlotPlan中的CORCIssue
        /// </summary>
        public void UpdateSlotCORC()
        {
            List<SlotCORC> slotCORC = GetSlotCORC();
            foreach (SlotCORC item in slotCORC)
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
                SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                if (sp != null)
                {
                    string update = "UPDATE KANBAN_SLOTPLAN SET CORC = @CORC, CORC_Issue = @CORC_Issue WHERE Slot = @Slot";
                    conn.Execute(update, item);
                }
            }
        }

        /// <summary>
        /// 获取Shortage
        /// </summary>
        public List<SlotShortage> GetSlotShortage()
        {
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\CSO";
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
                int week = Tool.WeekOfYear(DateTime.Now); // 获取当前周别
                // Type = P  Shortage != 0  Slot not contain Misc
                if (item[dt.Columns.IndexOf("Type")].ToString() == "P"
                    && item[dt.Columns.IndexOf("Shortage")].ToString() != "0"
                    && item[dt.Columns.IndexOf("Slot")].ToString().IndexOf("Misc") == -1
                    && (item[dt.Columns.IndexOf("WK")].ToString() == "WK" + week || item[dt.Columns.IndexOf("WK")].ToString() == "WK" + (week + 1)))
                {
                    SlotShortage info = new SlotShortage();
                    info.Slot = item[dt.Columns.IndexOf("Slot")].ToString();
                    info.PN = item[dt.Columns.IndexOf("Component")].ToString();
                    info.Description = item[dt.Columns.IndexOf("CompDescription")].ToString();
                    info.QTY = item[dt.Columns.IndexOf("ExtendedQty")].ToString();
                    info.ETA = item[dt.Columns.IndexOf("ConfirmedDate")].ToString();
                    info.IsReceived = false;
                    info.LastUpdatedTime = DateTime.Now;
                    list.Add(info);
                }
            }
            return list;
        }

        /// <summary>
        /// 插入或更新Shortage
        /// </summary>
        /// <param name="slotshortage"></param>
        public void InsertOrUpdateSlotShortage()
        {
            List<SlotShortage> slotshortage = GetSlotShortage();
            foreach (SlotShortage item in slotshortage)
            {
                string query = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE SLOT = @SLOT AND PN = @PN";
                List<SlotShortage> list = conn.Query<SlotShortage>(query, item).ToList();
                if (list.Count > 0)
                {
                    string update = "UPDATE KANBAN_SLOTSHORTAGE SET QTY = @QTY, LastUpdatedTime = @LastUpdatedTime WHERE SLOT = @SLOT AND PN = @PN";
                    conn.Execute(update, item);
                }
                else
                {
                    string insert = "INSERT INTO KANBAN_SLOTSHORTAGE VALUES (@SLOT, @PN, @Description, @QTY, @ETA, @IsReceived, @LastUpdatedTime)";
                    conn.Execute(insert, item);
                }
            }
        }

        /// <summary>
        /// 获取Config
        /// </summary>
        public List<SlotConfig> GetSlotConfig()
        {
            string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
            List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList();
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
                "TDN-636-860-21",
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
            string[] ZeroPin_List = new string[] {
                // UF 24
                "TDN-621-987-00",
                "TDN-621-987-01",
                // UF 12
                "TDN-979-712-00",
                // DRAGON 24
                "TDN-633-724-00",
                // DRAGON 12
                "TDN-633-724-12",
                // MF
                "TDN-202-000-01"
            };
            List<SlotConfig> list = new List<SlotConfig>();
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\E-Slot";
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
                if (slotplan.Find(x => x.Slot == item["Slot"].ToString()) != null && PN_List.Contains(item["Component"].ToString()))
                {
                    SlotConfig info = new SlotConfig();
                    info.Slot = item["Slot"].ToString();
                    info.PN = item["Component"].ToString().Substring(4);
                    info.QTY = item["Extended Qty"].ToString();
                    info.DelayTips = "";
                    info.IsReady = false;
                    info.LastUpdatedTime = DateTime.Now;
                    list.Add(info);
                }
                if (slotplan.Find(x => x.Slot == item["Slot"].ToString()) != null && ZeroPin_List.Contains(item["Component"].ToString()))
                {
                    SlotConfig info = new SlotConfig();
                    info.Slot = item["Slot"].ToString();
                    info.PN = "0-Pin";
                    info.QTY = item["Extended Qty"].ToString();
                    info.DelayTips = "";
                    info.IsReady = false;
                    info.LastUpdatedTime = DateTime.Now;
                    list.Add(info);
                }
            }
            return list;
        }

        /// <summary>
        /// 插入或更新config
        /// </summary>
        /// <param name="slotconfig"></param>
        public void InsertOrUpdateSlotConfig()
        {
            List<SlotConfig> slotconfig = GetSlotConfig();
            foreach (SlotConfig item in slotconfig)
            {
                string query = "SELECT * FROM KANBAN_SLOTCONFIG WHERE SLOT = @SLOT AND PN = @PN";
                List<SlotShortage> list = conn.Query<SlotShortage>(query, item).ToList();
                if (list.Count > 0)
                {
                    string update = "UPDATE KANBAN_SLOTCONFIG SET QTY = @QTY, LastUpdatedTime = @LastUpdatedTime WHERE SLOT = @SLOT AND PN = @PN";
                    conn.Execute(update, item);
                }
                else
                {
                    string insert = "INSERT INTO KANBAN_SLOTCONFIG VALUES (@SLOT, @PN, @Description, @QTY, @DelayTips, @IsReady, @LastUpdatedTime)";
                    conn.Execute(insert, item);
                }
            }
        }

        /// <summary>
        /// 更新Station Status
        /// </summary>
        /// <param name="slotconfig"></param>
        public void UpdateSlotStatus()
        {
            string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
            List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList().Where(x => GetValidate(x.MRP, 2) && x.Type == "UF").ToList();
            foreach (SlotPlan item in slotplan)
            {
                SlotStatus status = new SlotStatus();
                if (item.Model.IndexOf("24") > -1)
                {
                    status = GetUF24Status(item.Slot);
                }
                if (item.Model.IndexOf("12") > -1)
                {
                    status = GetUF12Status(item.Slot);
                }
                item.CoreBU = status.CorBU;
                item.PV = status.PV;
                item.OI = status.OI;
                item.TestBU = status.TestBU;
                item.CSW = status.CSW;
                item.QFAA = status.QFAA;
                item.BU = status.BU;
                item.Pack = status.Pack;
                string update = "UPDATE KANBAN_SLOTPLAN SET CoreBU = @CoreBU, PV = @PV, OI = @OI, TestBU = @TestBU, CSW = @CSW, QFAA = @QFAA, BU = @BU, Pack = @Pack WHERE ID = @ID";
                conn.Execute(update, item);
            }
        }

        /// <summary>
        /// 更新是否在minione
        /// </summary>
        /// <param name="slotconfig"></param>
        public void UpdateSlotMinioneStatus()
        {
            string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
            List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList();
            foreach (SlotPlan item in slotplan)
            {
                string sql = "SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Boards] WHERE TESTERID = @TESTERID";
                MinioneConfig mc = new MinioneConfig();
                mc.TesterID = item.Slot;
                List<MinioneConfig> MinioneList = connMinione.Query<MinioneConfig>(sql, mc).ToList();
                if (MinioneList.Count > 0)
                {
                    item.IsLoad = true;
                    item.LoadInfo = "";
                    foreach (MinioneConfig config in MinioneList)
                    {
                        item.LoadInfo = item.LoadInfo + "," + config.PN;
                    }
                    item.LoadInfo = item.LoadInfo.Substring(1);
                }
                else
                {
                    string sqlChecklist = "SELECT * FROM [SC3_test].[dbo].[SystemProperties] WHERE systemName = '" + item.Slot + "'";
                    ChecklistSystemInfo checklistSystemInfo = conn.Query<ChecklistSystemInfo>(sqlChecklist).SingleOrDefault();
                    if (checklistSystemInfo != null)
                    {
                        mc = new MinioneConfig();
                        mc.TesterID = checklistSystemInfo.systemId.Substring(1, 3) + checklistSystemInfo.systemId.Substring(5, 3);
                        MinioneList = connMinione.Query<MinioneConfig>(sql, mc).ToList();
                        if (MinioneList.Count > 0)
                        {
                            item.IsLoad = true;
                            item.LoadInfo = "";
                            foreach (MinioneConfig config in MinioneList)
                            {
                                item.LoadInfo = item.LoadInfo + "," + config.PN;
                            }
                            item.LoadInfo = item.LoadInfo.Substring(1);
                        }
                        else
                        {
                            item.IsLoad = false;
                            item.LoadInfo = "";
                        }
                    }
                    else
                    {
                        item.IsLoad = false;
                        item.LoadInfo = "";
                    }
                }
                string update = "UPDATE KANBAN_SLOTPLAN SET IsLoad = @IsLoad, LoadInfo = @LoadInfo WHERE ID = @ID";
                conn.Execute(update, item);
            }
        }

        /// <summary>
        /// 更新PSB SSB在系统中的数量
        /// </summary>
        public void UpdatePSBSSBQty()
        {
            PSBSSB obj = new PSBSSB();
            obj.PSBQty = GetConfigQty("TDN-974-221-20");
            obj.SSBQty = GetConfigQty("TDN-974-221-30");
            obj.LastUpdatedTime = DateTime.Now;
            string sql = "Update KANBAN_PSBSSB set PSBQty = @PSBQty, SSBQty = @SSBQty, LastUpdatedTime = @LastUpdatedTime";
            conn.Execute(sql, obj);
        }





        /// <summary>
        /// 获取系统的单个状态的方法
        /// </summary>
        /// <param name="slot"></param>
        /// <param name="station"></param>
        /// <returns></returns>
        public string GetStatusByItem(string slot, string item, string station, string tableName)
        {
            string sql = @"SELECT TOP 1 SUBSTRING(CONVERT(varchar, DATENAME(YEAR,Date)),3,2) + " +
                         "RIGHT('00' + CAST(DATENAME(WEEK, a.Date) AS nvarchar(50)), 2) + '.' + " +
                         "case when convert(varchar(50), datepart(dw, A.Date) - 1) = '0' then '7' else convert(varchar(50), datepart(dw, A.Date) - 1) end as WeekNum, b.systemName as Slot " +
                         "FROM [SC3_test].[dbo].[" + tableName + "] a " +
                         "join [SC3_test].[dbo].[SystemProperties] b " +
                         "on a.systemId = b.systemId " +
                         "where b.systemName = '" + slot + "' and a.CheckItemDescription like '%" + item + "%'" + " and a.ListBy like '%" + station + "%' " +
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
        public SlotStatus GetUF24Status(string slot)
        {
            SlotStatus status = new SlotStatus();
            status.Slot = slot;
            status.Launch = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex SC3 Core Bring Up Station Checklist", "Check_SC3");
            status.CorBU = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex SC3 Core Bring Up Station Checklist", "Check_SC3");
            status.PV = GetStatusByItem(slot, "1. QDC Latch(inside test head)", "Ultraflex SC3 Core Bring Up Station Checklist", "Check_SC3");
            status.OI = GetStatusByItem(slot, "33. Clamp(To heat exchange)", "Ultraflex SC3 Core Bring Up Station Checklist", "Check_SC3");
            status.TestBU = GetStatusByItem(slot, @"1.1 Verify latest PO and CORC.Save CORC and Slot Locator in below path: \\ Agate\UltraFlex Project\Sync data\Shipped\CORC Tracking", "Ultraflex SC3 Test Bring Up Station Checklist", "Check_SC3");
            status.CSW = GetStatusByItem(slot, "Verify limit table and firmware updated before CSW", "Ultraflex SC3 Test Bring Up Station Checklist", "Check_SC3");
            status.QFAA = GetStatusByItem(slot, "1.1 Record HFE Level", "Ultraflex SC3 Final Station Checklist", "Check_SC3");
            status.BU = GetStatusByItem(slot, "1.1打开Test Head 的门,确认门两侧各有2个Danger 标签且无缺损脏污.", "Ultraflex SC3 Assembly Button up Checklist", "Check_SC3");
            status.Pack = GetStatusByItem(slot, "1. 木箱检查：检查没有多余文字或者图，没有大面积划伤（图1）；检查没有多余的毛刺毛边杂物，如果有需要清理干净（图2,3,4）。检查木箱所有钉子已经钉好，没有外漏，如果有需要拔出来重新钉（图5,6,7）。", "Ultraflex SC3 Packing Checklist", "Check_SC3");
            return status;
        }

        /// <summary>
        /// 获取系统整个状态的对象
        /// </summary>
        /// <param name="slot"></param>
        /// <returns></returns>
        public SlotStatus GetUF12Status(string slot)
        {
            SlotStatus status = new SlotStatus();
            status.Slot = slot;
            status.Launch = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", "Check_12Slot");
            status.CorBU = GetStatusByItem(slot, "1.1 Please set IP address according to the tag on the desk", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", "Check_12Slot");
            status.PV = GetStatusByItem(slot, "1. QDC Latch(inside test head)", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", "Check_12Slot");
            status.OI = GetStatusByItem(slot, "25. TH pipe drain QD(left)", "Ultraflex 12Slot Core Bring Up Verification Station Checklist", "Check_12Slot");
            status.TestBU = GetStatusByItem(slot, @"1.2 Verify latest PO and CorC. Save CORC and Slot Locator in below path: \\ Agate\UltraFlex Project\Sync data\Shipped\CORC Tracking", "Ultraflex 12Slot Test Bring Up Station Checklist", "Check_12Slot");
            status.CSW = GetStatusByItem(slot, "1.29 Verify limit table and firmware updated before CSW", "Ultraflex 12Slot Test Bring Up Station Checklist", "Check_12Slot");
            status.QFAA = GetStatusByItem(slot, "1.1 Record HFE Level", "UltraFlex 12Slot Final Station Checklist", "Check_12Slot");
            status.BU = GetStatusByItem(slot, "1.1打开Test Head 的门,确认门侧有2个Dangerous 标签且无缺损脏污", "Ultraflex 12Slot Assembly Button up Checklist", "Check_12Slot");
            status.Pack = GetStatusByItem(slot, "1. 木箱检查：检查没有多余文字或者图，没有大面积划伤（图1）；检查没有多余的毛刺毛边杂物，如果有需要清理干净（图2,3,4）。检查木箱所有钉子已经钉好，没有外漏，如果有需要拔出来重新钉（图5,6,7）。", "Ultraflex 12Slot Packing Checklist", "Check_12Slot");
            return status;
        }

        public bool GetValidate(string mrpStr)
        {
            int week = Tool.WeekOfYear(DateTime.Now); // 获取当周周别
            float mrp = float.Parse(mrpStr); // 获取MRP
            int range = 5; // 抓取的周范围 后期改成可修改
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

        public bool GetValidate(string mrpStr, int range)
        {
            int week = Tool.WeekOfYear(DateTime.Now); // 获取当周周别
            float mrp = float.Parse(mrpStr); // 获取MRP
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

        public int GetConfigQty(string pn)
        {
            int qty = 0;
            List<SlotConfig> list = new List<SlotConfig>();
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\E-Slot";
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
                if ((item["Slot"].ToString() == "FCT burning 24sl" || item["Slot"].ToString() == "FCT burning SC3.0" || item["Slot"].ToString() == "FCT verify 12sl") && item["Component"].ToString() == pn)
                {
                    qty += int.Parse(item["Extended Qty"].ToString());
                }
            }
            return qty;
        }
    }
}