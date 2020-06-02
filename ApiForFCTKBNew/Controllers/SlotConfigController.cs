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
    public class SlotConfigController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public string UpdateSlotConfig()
        {
            try
            {
                // 查询出未出货的系统
                string query = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate is null";
                List<SlotPlan> slotplan = conn.Query<SlotPlan>(query).ToList();

                // 定义需要匹配的板号
                string[] PN_List = new string[] {
                "TDN-974-221-20",
                "TDN-974-221-30",
                "TDN-604-356-01",
                "TDN-604-356-03",
                "TDN-604-356-04",
                "TDN-805-052-02",
                "TDN-805-052-05",
                "TDN-974-294-30",
                "TDN-604-375-02",
                "TDN-605-743-01",
                "TDN-974-217-30",
                "TDN-974-390-02",
                "TDN-974-230-00",
                "TDN-974-232-00",
                "TDN-630-036-30",
                "TDN-636-860-21",
                "TDN-974-296-30",
                "TDN-632-860-21",
                "TDN-633-675-03",
                "TDN-630-035-40",
                "TDN-617-743-00",
                "TDN-617-744-00",
                "TDN-974-245-40",
                "TDN-604-375-12",
                "TDN-625-393-51",
                "TDN-639-209-02",
                "TDN-654-990-00",
                "TDN-805-002-60",
                "TDN-805-003-11",
                "TDN-805-003-50",
                "TDN-805-229-61",
                "TDN-805-230-60",
                "TDN-805-251-50",
                "TDN-805-333-50",
                "TDN-805-386-01",
                "TDN-805-740-50",
                "TDN-949-974-50",
                // J750
                "TDN-615-988-01",
                "TDN-614-889-01",
                "TDN-805-025-04",
                "TDN-239-701-00",
                "TDN-239-016-06",
                "TDN-314-445-51",
                "TDN-239-029-02",
                "TDN-616-451-01",
                "TDN-517-440-00",
                "TDN-314-446-03",
                "TDN-239-046-93",
                "TDN-623-717-00",
                "TDN-617-084-01",
                "TDN-239-004-03",
                "TDN-239-052-03",
                "TDN-517-456-00",
                "TDN-624-102-01",

                "TDN-239-046-93-P",
                "TDN-239-052-03-M",
                "TDN-617-084-01-M",
                "TDN-624-102-01-M"

            };
                string[] ZeroPin_List = new string[] {
                // UF 36
                "TDN-617-097-00",
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

                List<SlotConfig> slotconfig = new List<SlotConfig>();

                //读取Excel
                string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\E-Slot";
                DirectoryInfo dir = new DirectoryInfo(path);
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    if (file.FullName.ToUpper().Contains("SLOT"))
                    {
                        path = file.FullName;
                        break;
                    }
                }
                Workbook wb = new Workbook(path);
                Cells cells = wb.Worksheets[0].Cells;
                DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);

                // PSB SSB 数量
                PSBSSB obj = new PSBSSB();

                // 遍历数据，找到需要的板子和0-pin
                foreach (DataRow item in dt.Rows)
                {
                    // PSB SSB 数量
                    if ((item["Slot"].ToString() == "FCT burning 24sl" || item["Slot"].ToString() == "FCT burning SC3.0" || item["Slot"].ToString() == "FCT verify 12sl") && item["Component"].ToString() == "TDN-974-221-20")
                    {
                        obj.PSBQty += int.Parse(item["Extended Qty"].ToString());
                    }

                    if ((item["Slot"].ToString() == "FCT burning 24sl" || item["Slot"].ToString() == "FCT burning SC3.0" || item["Slot"].ToString() == "FCT verify 12sl") && item["Component"].ToString() == "TDN-974-221-30")
                    {
                        obj.SSBQty += int.Parse(item["Extended Qty"].ToString());
                    }

                    SlotPlan sp = slotplan.Where(x => x.Slot == item["Slot"].ToString()).FirstOrDefault();
                    if (sp != null && PN_List.Contains(item["Component"].ToString()))
                    {
                        SlotConfig info = new SlotConfig();
                        info.Slot = item["Slot"].ToString();
                        info.PN = item["Component"].ToString().Substring(4);
                        info.QTY = int.Parse(item["Extended Qty"].ToString());
                        info.DelayTips = "";
                        info.IsReady = 0;
                        info.LastUpdatedTime = DateTime.Now;
                        slotconfig.Add(info);
                    }
                    if (sp != null && ZeroPin_List.Contains(item["Component"].ToString()))
                    {
                        SlotConfig info = new SlotConfig();
                        info.Slot = item["Slot"].ToString();
                        info.PN = "0-Pin";
                        info.QTY = int.Parse(item["Extended Qty"].ToString());
                        info.DelayTips = "";
                        info.IsReady = 0;
                        info.LastUpdatedTime = DateTime.Now;
                        slotconfig.Add(info);
                    }
                }

                obj.LastUpdatedTime = DateTime.Now;
                string sql = "Update KANBAN_PSBSSB set PSBQty = @PSBQty, SSBQty = @SSBQty, LastUpdatedTime = @LastUpdatedTime";
                conn.Execute(sql, obj);


                //// 获取数据库的config信息
                //string sqlQuery = "SELECT a.* FROM KANBAN_SLOTCONFIG a JOIN KANBAN_SLOTPLAN b on a.Slot = b.Slot WHERE b.ShippingDate IS NULL";
                //List<SlotConfig> listDB = conn.Query<SlotConfig>(sqlQuery).ToList();

                //// 删除老的 添加新的
                //string delSql = "DELETE KANBAN_SLOTCONFIG WHERE ID = @ID";
                //conn.Execute(delSql, listDB);

                //string insertSql = "INSERT INTO KANBAN_SLOTCONFIG VALUES (@SLOT, @PN, @Description, @QTY, @DelayTips, @IsReady, @LastUpdatedTime)";
                //conn.Execute(insertSql, slotconfig);

                foreach (SlotConfig item in slotconfig)
                {
                    string querySQL = "SELECT * FROM KANBAN_SLOTCONFIG WHERE SLOT = @SLOT AND PN = @PN";
                    List<SlotConfig> list = conn.Query<SlotConfig>(querySQL, item).ToList();
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

                string qSQL = "SELECT a.* FROM [SlotKB].[dbo].[KANBAN_SLOTCONFIG] a join [SlotKB].[dbo].[KANBAN_SLOTPLAN] b on a.Slot = b.Slot where b.ShippingDate is null";
                List<SlotConfig> SlotConfigList = conn.Query<SlotConfig>(qSQL).ToList();
                foreach (SlotConfig item in SlotConfigList)
                {
                    bool flag = true;
                    foreach (SlotConfig i in slotconfig)
                    {
                        if (item.Slot == i.Slot && item.PN == i.PN)
                        {
                            flag = false;
                        }
                    }
                    if (flag)
                    {
                        string delSql = "DELETE KANBAN_SLOTCONFIG WHERE SLOT = @SLOT AND PN = @PN";
                        conn.Execute(delSql, item);
                    }
                }

                return "E-SLOT OK!";
            }
            catch (Exception ex)
            {
                return "E-SLOT Failed! " + ex.Message;
            }
        }

        [HttpPost]
        public int UpdateConfig(SlotConfig config)
        {
            string update = "UPDATE KANBAN_SLOTCONFIG SET DelayTips = @DelayTips, IsReady = @IsReady WHERE ID = @ID";
            return conn.Execute(update, config);
        }

        private class PSBSSB
        {
            public int PSBQty { get; set; }
            public int SSBQty { get; set; }
            public DateTime LastUpdatedTime { get; set; }
        }
    }
}
