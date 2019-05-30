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
    public class ConfigController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpPost]
        public string UpdateConfig(SlotConfig config)
        {
            string update = "UPDATE KANBAN_SLOTCONFIG SET DelayTips = @DelayTips, IsReady = @IsReady WHERE ID = @ID";
            conn.Execute(update, config);
            return "ok";
        }
    }
}
