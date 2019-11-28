using ActionTracker.Models;
using Aspose.Cells;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ActionTracker.Controllers
{
    public class ConfigController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        [HttpGet]
        public void GetSlotPlan()
        {
            string path = @"\\10.194.51.14\Backup\Jesse\WW_IB_TSMC.xlsx";
            List<Config> list = new List<Config>();
            Workbook wb = new Workbook(path);
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            string sntotal = "";
            foreach (DataRow item in dt.Rows)
            {
                string systemSN = item[dt.Columns.IndexOf("SystemSN")].ToString();
                if (sntotal == "")
                {
                    sntotal = "'" + systemSN + "'";
                }
                else
                {
                    sntotal = sntotal + "," + "'" + systemSN + "'";
                }
            }
            string sqlStr = @"select a.TesterSN, a.TesterID, b.SN, b.Slot from [172.21.194.214].[PCBA].[dbo].[SystemInfo] a 
                            join [172.21.194.214].[PCBA].[dbo].[Boards] b
                            on a.TesterID = b.TesterID
                            where a.TesterSN in (" + sntotal + ")";

            //string sql = "select distinct TesterSN from [172.21.194.214].[PCBA].[dbo].[SystemInfo] where TesterSN in (" + sntotal + ") order by TesterSN";
            List<SystemSlot> systemSlots = conn.Query<SystemSlot>(sqlStr).ToList();

            foreach (DataRow item in dt.Rows)
            {
                List<SystemSlot> temp = systemSlots.Where(x => x.TesterSN == item[dt.Columns.IndexOf("SystemSN")].ToString()).ToList();
                SystemSlot systemSlot = temp.Count > 0 ? temp[0] : null;
                if (systemSlot != null)
                {
                    Config config = new Config();
                    config.SystemSN = systemSlot.TesterSN;
                    config.SystemName = systemSlot.TesterID;
                    config.Slot0 = GetSystemSlot(systemSlots, "0", dt, item);
                    config.Slot1 = GetSystemSlot(systemSlots, "1", dt, item);
                    config.Slot2 = GetSystemSlot(systemSlots, "2", dt, item);
                    config.Slot3 = GetSystemSlot(systemSlots, "3", dt, item);
                    config.Slot4 = GetSystemSlot(systemSlots, "4", dt, item);
                    config.Slot5 = GetSystemSlot(systemSlots, "5", dt, item);
                    config.Slot6 = GetSystemSlot(systemSlots, "6", dt, item);
                    config.Slot7 = GetSystemSlot(systemSlots, "7", dt, item);
                    config.Slot8 = GetSystemSlot(systemSlots, "8", dt, item);
                    config.Slot9 = GetSystemSlot(systemSlots, "9", dt, item);
                    config.Slot10 = GetSystemSlot(systemSlots, "10", dt, item);
                    config.Slot11 = GetSystemSlot(systemSlots, "11", dt, item);
                    config.Slot12 = GetSystemSlot(systemSlots, "12", dt, item);
                    config.Slot13 = GetSystemSlot(systemSlots, "13", dt, item);
                    config.Slot14 = GetSystemSlot(systemSlots, "14", dt, item);
                    config.Slot15 = GetSystemSlot(systemSlots, "15", dt, item);
                    config.Slot16 = GetSystemSlot(systemSlots, "16", dt, item);
                    config.Slot17 = GetSystemSlot(systemSlots, "17", dt, item);
                    config.Slot18 = GetSystemSlot(systemSlots, "18", dt, item);
                    config.Slot19 = GetSystemSlot(systemSlots, "19", dt, item);
                    config.Slot20 = GetSystemSlot(systemSlots, "20", dt, item);
                    config.Slot21 = GetSystemSlot(systemSlots, "21", dt, item);
                    config.Slot22 = GetSystemSlot(systemSlots, "22", dt, item);
                    config.Slot23 = GetSystemSlot(systemSlots, "23", dt, item);
                    config.Slot24 = GetSystemSlot(systemSlots, "24", dt, item);
                    config.Slot25 = GetSystemSlot(systemSlots, "25", dt, item);
                    list.Add(config);
                }
            }

            for (int i = 1; i < cells.MaxDataRow + 1; i++)
            {
                string SystemSN = cells[i, 0].Value.ToString();
                Config config = list.Where(x => x.SystemSN == SystemSN).SingleOrDefault();
                if (config != null)
                {
                    cells[i, 1].PutValue(config.SystemName);
                    cells[i, 2].PutValue(config.Slot24);
                    cells[i, 3].PutValue(config.Slot25);
                    cells[i, 4].PutValue(config.Slot0);
                    cells[i, 5].PutValue(config.Slot1);
                    cells[i, 6].PutValue(config.Slot2);
                    cells[i, 7].PutValue(config.Slot3);
                    cells[i, 8].PutValue(config.Slot4);
                    cells[i, 9].PutValue(config.Slot5);
                    cells[i, 10].PutValue(config.Slot6);
                    cells[i, 11].PutValue(config.Slot7);
                    cells[i, 12].PutValue(config.Slot8);
                    cells[i, 13].PutValue(config.Slot9);
                    cells[i, 14].PutValue(config.Slot10);
                    cells[i, 15].PutValue(config.Slot11);
                    cells[i, 16].PutValue(config.Slot12);
                    cells[i, 17].PutValue(config.Slot13);
                    cells[i, 18].PutValue(config.Slot14);
                    cells[i, 19].PutValue(config.Slot15);
                    cells[i, 20].PutValue(config.Slot16);
                    cells[i, 21].PutValue(config.Slot17);
                    cells[i, 22].PutValue(config.Slot18);
                    cells[i, 23].PutValue(config.Slot19);
                    cells[i, 24].PutValue(config.Slot20);
                    cells[i, 25].PutValue(config.Slot21);
                    cells[i, 26].PutValue(config.Slot22);
                    cells[i, 27].PutValue(config.Slot23);
                }
            }

            wb.Save(@"\\10.194.51.14\Backup\Jesse\1.xlsx");
        }

        public string GetSystemSlot(List<SystemSlot> list, String Slot, DataTable dt, DataRow item)
        {
            List<SystemSlot> temp = list.Where(x => x.TesterSN == item[dt.Columns.IndexOf("SystemSN")].ToString()).ToList();
            SystemSlot systemSlot = temp.Count > 0 ? temp[0] : null;
            return list.Where(x => (x.TesterSN == item[dt.Columns.IndexOf("SystemSN")].ToString()
            && x.TesterID == systemSlot.TesterID && x.Slot == Slot)).SingleOrDefault() == null ? "N/A" :
            list.Where(x => (x.TesterSN == item[dt.Columns.IndexOf("SystemSN")].ToString()
            && x.TesterID == systemSlot.TesterID && x.Slot == Slot)).SingleOrDefault().SN;
        }
    }
}
