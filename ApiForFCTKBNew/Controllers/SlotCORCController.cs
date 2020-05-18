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
    public class SlotCORCController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        public string UpdateSlotCORC()
        {
            try
            {
                // 更新slotplan中的CORC字段
                List<SlotCORC> slotCORC = GetSlotCORC();
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
                foreach (SlotCORC item in slotCORC)
                {
                    SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    if (sp != null)
                    {
                        string update = "UPDATE KANBAN_SLOTPLAN SET CORC = @CORC, CORC_Issue = @CORC_Issue WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                }

                // 更新kanban_slotcorc表
                // 遍历数据库closetime为空的数据， 如果excel中查不到，就默认closetime为当前时间
                string q = "SELECT * FROM KANBAN_SLOTCORC WHERE CloseTime is null and Slot <> ''";
                List<SlotCORC> list = conn.Query<SlotCORC>(q).ToList();

                foreach (SlotCORC item in list)
                {
                    SlotCORC sp = slotCORC.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    if (sp == null)
                    {
                        item.CloseTime = DateTime.Now;
                        string update = "UPDATE KANBAN_SLOTCORC SET CORC_Issue = @CORC_Issue, CloseTime = @CloseTime WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                }

                list = conn.Query<SlotCORC>(q).ToList();

                foreach (SlotCORC item in slotCORC)
                {
                    SlotCORC sp = list.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    // 有记录 CORC 不为空
                    if (sp != null && item.CORC_Issue != "" && item.CORC_Issue != null)
                    {
                        if (sp.CORC_Issue != item.CORC_Issue)
                        {
                            item.LastUpdateTime = DateTime.Now;
                            string update = "UPDATE KANBAN_SLOTCORC SET CORC_Issue = @CORC_Issue, LastUpdateTime = @LastUpdateTime WHERE Slot = @Slot";
                            conn.Execute(update, item);
                        }
                    }
                    // 有记录 CORC 为空
                    else if (sp != null && !(item.CORC_Issue != "" && item.CORC_Issue != null))
                    {
                        item.CloseTime = DateTime.Now;
                        item.LastUpdateTime = DateTime.Now;
                        string update = "UPDATE KANBAN_SLOTCORC SET LastUpdateTime = @LastUpdateTime, CloseTime = @CloseTime WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                    // 无记录 CORC 不为空
                    else if (sp == null && item.CORC_Issue != "" && item.CORC_Issue != null)
                    {
                        item.InsertTime = DateTime.Now;
                        string insert = "INSERT INTO KANBAN_SLOTCORC (Slot, CORC_Issue, InsertTime) VALUES (@Slot, @CORC_Issue, @InsertTime)";
                        conn.Execute(insert, item);
                    }
                    // 无记录 CORC为空
                    else
                    {
                        // Nothing to do
                    }
                }
                return "SlotCORC OK!";
            }
            catch (Exception ex)
            {
                return "SlotCORC Failed! " + ex.Message;
            }
        }

        public List<SlotCORC> GetSlotCORC()
        {
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\SO&CORC Issue Report";
            string flexPath = "";
            string JPath = "";
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
                    flexPath = file.FullName;
                }
                if (file.FullName.Contains("J750"))
                {
                    JPath = file.FullName;
                }
            }
            List<SlotCORC> list = new List<SlotCORC>();
            Workbook wb = new Workbook(flexPath);
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

            // J750
            if (JPath != "")
            {
                wb = new Workbook(JPath);
                foreach (Worksheet ws in wb.Worksheets)
                {
                    if (ws.Name.IndexOf("Oline System") > -1)
                    {
                        Cells cells = ws.Cells;
                        DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                        foreach (DataRow item in dt.Rows)
                        {
                            SlotCORC info = new SlotCORC();
                            info.Slot = item[dt.Columns.IndexOf("SLOT")].ToString();
                            string status = item[dt.Columns.IndexOf("CORC Status")].ToString();
                            info.CORC_Issue = item[dt.Columns.IndexOf("SO or CORC issue")].ToString();
                            if (status == "IN REVIEW")
                            {
                                info.CORC = "R";
                            }
                            else if (status == "NO SO&CORC")
                            {
                                info.CORC = "";
                                info.CORC_Issue = "NO SO&CORC";
                            }
                            else if (status.Contains("Approved"))
                            {
                                info.CORC = "A";
                            }
                            list.Add(info);
                        }
                    }
                }
            }
            return list;
        }

        public class SlotCORC
        {
            public string Slot { get; set; }
            public string CORC { get; set; }
            public string CORC_Issue { get; set; }
            public DateTime InsertTime { get; set; }
            public DateTime LastUpdateTime { get; set; }
            public DateTime CloseTime { get; set; }
        }
    }
}
