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
    public class SlotPlanController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpPost]
        public int UpdateSlotPlan (SlotPlan slotplan)
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
    }
}
