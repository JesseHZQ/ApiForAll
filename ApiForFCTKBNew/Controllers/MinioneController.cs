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
    public class MinioneController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        [HttpGet]
        public string UpdateSlotMinioneStatus()
        {
            try
            {
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList();

                string sql = "SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Boards]";
                List<MinioneConfig> MinioneList = connMinione.Query<MinioneConfig>(sql).ToList();

                string sqlChecklist = "SELECT * FROM [SC3_test].[dbo].[SystemProperties]";
                List<ChecklistSystemInfo> csList = conn.Query<ChecklistSystemInfo>(sqlChecklist).ToList();

                foreach (SlotPlan item in slotplan)
                {
                    // 截取数字部分 2019/12/20 修改
                    if (item.Slot.IndexOf("D") > -1 && item.Slot.IndexOf("T") == -1 && item.Slot.IndexOf("A") == -1)
                    {
                        string num = item.Slot.IndexOf("B") > -1 ? item.Slot.Substring(1, item.Slot.Length - 2) : item.Slot.Substring(1, item.Slot.Length - 1);
                        num = num.Length == 2 ? "00" + num : (num.Length == 3 ? "0" + num : num);
                        item.Slot = item.Slot.IndexOf("B") > -1 ? "D" + num + "B" : "D" + num + "A";
                    }

                    List<MinioneConfig> list = MinioneList.Where(x => x.TesterID == item.Slot).ToList();

                    if (list.Count > 0)
                    {
                        item.IsLoad = true;
                        item.LoadInfo = "";
                        foreach (MinioneConfig config in list)
                        {
                            item.LoadInfo = item.LoadInfo + "," + config.PN;
                        }
                        item.LoadInfo = item.LoadInfo.Substring(1);
                    }
                    else
                    {
                        ChecklistSystemInfo csi = csList.Where(x => x.systemName == item.Slot).FirstOrDefault();
                        if (csi != null)
                        {
                            string testerID = csi.systemId.Substring(1, 3) + csi.systemId.Substring(5, 3);
                            list = MinioneList.Where(x => x.TesterID == testerID).ToList();
                            if (list.Count > 0)
                            {
                                item.IsLoad = true;
                                item.LoadInfo = "";
                                foreach (MinioneConfig config in list)
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
                return "Minione OK!";
            }
            catch (Exception ex)
            {
                return "Minione Failed! " + ex.Message;
            }
        }

        private class MinioneConfig
        {
            public string SN { get; set; }
            public string PN { get; set; }
            public string TesterID { get; set; }
        }

        private class ChecklistSystemInfo
        {
            public string systemId { get; set; }
            public string systemName { get; set; }
        }
    }
}
