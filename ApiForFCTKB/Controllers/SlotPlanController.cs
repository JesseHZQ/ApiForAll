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
        public int DeleteSlotPlan(SlotPlan slotplan)
        {
            string sql = "DELETE KANBAN_SLOTPLAN WHERE ID = @ID";
            return conn.Execute(sql, slotplan);
        }
    }
}
