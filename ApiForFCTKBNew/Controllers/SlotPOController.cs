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
    public class SlotPOController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        public List<SlotPO> GetSlotPO()
        {
            // 读取open_po
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\Open_PO";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            string flexPath = "";
            string J750Pth = "";
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("Open_PO"))
                {
                    flexPath = file.FullName;
                }
                if (file.FullName.Contains("J750"))
                {
                    J750Pth = file.FullName;
                }
            }

            // UF
            List<SlotPO> list = new List<SlotPO>();
            Workbook wb = new Workbook(flexPath);
            foreach (Worksheet ws in wb.Worksheets)
            {
                Cells cells = ws.Cells;
                DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
                foreach (DataRow item in dt.Rows)
                {
                    string slot = item[dt.Columns.IndexOf(@"Slot Number\Misc Order")].ToString();
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

            //J750
            wb = new Workbook(J750Pth);
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
        
        [HttpGet]
        public string UpdateSlotPO()
        {
            try
            {
                List<SlotPO> slotpo = GetSlotPO();
                string query = "SELECT * FROM KANBAN_SLOTPLAN";
                List<SlotPlan> slotlist = conn.Query<SlotPlan>(query).ToList();
                foreach (SlotPO item in slotpo)
                {
                    SlotPlan sp = slotlist.Where(x => x.Slot == item.Slot).SingleOrDefault();
                    if (sp != null)
                    {
                        string update = "UPDATE KANBAN_SLOTPLAN SET PO = @PO WHERE Slot = @Slot";
                        conn.Execute(update, item);
                    }
                }
                return "SlotPO OK!";
            }
            catch (Exception ex)
            {
                return "SlotPO Failed! " + ex.Message;
            }
        }

        public class SlotPO
        {
            public string Slot { get; set; }
            public string PO { get; set; }
        }
    }
}
