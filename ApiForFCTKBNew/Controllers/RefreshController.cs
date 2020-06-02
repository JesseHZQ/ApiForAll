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
    public class RefreshController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public string Refresh()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");

            string msg = "";
            SlotPlanController slotplan = new SlotPlanController();
            msg += slotplan.UpdateSlotPlan();
            SlotPOController slotpo = new SlotPOController();
            msg += slotpo.UpdateSlotPO();
            SlotCORCController slotcorc = new SlotCORCController();
            msg += slotcorc.UpdateSlotCORC();
            SlotShortageController slotshortage = new SlotShortageController();
            msg += slotshortage.UpdateSlotShortage();
            SlotConfigController slotconfig = new SlotConfigController();
            msg += slotconfig.UpdateSlotConfig();
            MinioneController minione = new MinioneController();
            msg += minione.UpdateSlotMinioneStatus();
            SlotStatusController slotstatus = new SlotStatusController();
            msg += slotstatus.UpdateSlotStatus();
            ShippingController shipping = new ShippingController();
            msg += shipping.UpdateShippingStatus();
            FCCPController fccp = new FCCPController();
            msg += fccp.UpdateFCCP();

            KanbanLog kanbanLog = new KanbanLog();
            kanbanLog.LogInfo = msg;
            kanbanLog.Date = DateTime.Now;

            string sqlLog = "insert into KANBAN_LOG VALUES (@LogInfo, @Date)";
            conn.Execute(sqlLog, kanbanLog);

            return msg;
        }

        [HttpGet]
        public KanbanLog GetRefreshTime()
        {
            string sql = "select top 1 * from KANBAN_LOG order by Date desc";
            return conn.Query<KanbanLog>(sql).ToList().FirstOrDefault();
        }

        public class KanbanLog
        {
            public int Id { get; set; }
            public string LogInfo { get; set; }
            public DateTime Date { get; set; }
        }
    }
}
