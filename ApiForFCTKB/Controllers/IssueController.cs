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

namespace ApiForFCTKB.Controllers
{
    public class IssueController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpPost]
        public string UpdateEIssue(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET Engineering_Issue = @Engineering_Issue WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }

        [HttpPost]
        public string UpdateRemark(SlotPlan slot)
        {
            string update = "UPDATE KANBAN_SLOTPLAN SET Remark = @Remark WHERE ID = @ID";
            conn.Execute(update, slot);
            return "ok";
        }
    }
}
