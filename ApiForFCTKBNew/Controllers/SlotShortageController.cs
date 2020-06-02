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
    public class SlotShortageController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

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
                int nextWeek = Tool.WeekOfYear(DateTime.Now.AddDays(7));
                string nextWk = nextWeek.ToString().Length > 1 ? nextWeek.ToString() : ("0" + nextWeek.ToString());
                string wk = week.ToString().Length > 1 ? week.ToString() : ("0" + week.ToString());
                if (item[dt.Columns.IndexOf("Type")].ToString() == "P"
                    && item[dt.Columns.IndexOf("Shortage")].ToString().Trim() != "0"
                    && item[dt.Columns.IndexOf("Slot")].ToString().IndexOf("Misc") == -1
                    && (item[dt.Columns.IndexOf("WK")].ToString() == "WK" + wk || item[dt.Columns.IndexOf("WK")].ToString() == "WK" + nextWk)
                    )
                {
                    SlotShortage info = new SlotShortage();
                    info.Slot = item[dt.Columns.IndexOf("Slot")].ToString();
                    info.PN = item[dt.Columns.IndexOf("Component")].ToString();
                    info.Description = item[dt.Columns.IndexOf("CompDescription")].ToString();
                    info.SupplierName = item[dt.Columns.IndexOf("SupplierName")].ToString();
                    info.Buyer = item[dt.Columns.IndexOf("Buyer")].ToString();
                    info.QTY = item[dt.Columns.IndexOf("ExtendedQty")].ToString();
                    info.ETA = item[dt.Columns.IndexOf("ConfirmedDate")].ToString();
                    info.IsReceived = false;
                    list.Add(info);
                }
            }
            return list;
        }

        [HttpGet]
        public string UpdateSlotShortage()
        {
            try
            {
                List<SlotShortage> slotshortage = GetSlotShortage();
                string query = "SELECT * FROM KANBAN_SLOTSHORTAGE";
                List<SlotShortage> list = conn.Query<SlotShortage>(query).ToList();
                foreach (SlotShortage item in slotshortage)
                {
                    List<SlotShortage> ss = list.Where(x => x.Slot == item.Slot && x.PN == item.PN).ToList();
                    if (ss.Count > 0)
                    {
                        item.LastUpdatedTime = DateTime.Now;
                        string update = "UPDATE KANBAN_SLOTSHORTAGE SET Buyer = @Buyer, SupplierName = @SupplierName, QTY = @QTY, LastUpdatedTime = @LastUpdatedTime WHERE SLOT = @SLOT AND PN = @PN";
                        conn.Execute(update, item);
                    }
                    else
                    {
                        item.InsertTime = DateTime.Now;
                        string insert = "INSERT INTO KANBAN_SLOTSHORTAGE (SLOT, PN, Description, SupplierName, QTY, Buyer, ETA, IsReceived, InsertTime) VALUES (@SLOT, @PN, @Description, @SupplierName, @QTY, @Buyer, @ETA, @IsReceived, @InsertTime)";
                        conn.Execute(insert, item);
                    }
                }
                return "SlotShortage OK!";
            }
            catch (Exception ex)
            {
                return "SlotShortage Failed! " + ex.Message;
            }
        }

        [HttpGet]
        public List<SlotShortage> GetShortageBySlot(string slot)
        {
            string sql = "SELECT * FROM KANBAN_SLOTSHORTAGE WHERE Slot = '" + slot + "'";
            return conn.Query<SlotShortage>(sql).ToList();
        }

        [HttpPost]
        public int UpdateShortage(SlotShortage shortage)
        {
            shortage.ReceivedTime = DateTime.Now;
            string update = "UPDATE KANBAN_SLOTSHORTAGE SET IsReceived = @IsReceived, ETA = @ETA, ReceivedTime = @ReceivedTime WHERE ID = @ID";
            return conn.Execute(update, shortage);
        }
    }
}
