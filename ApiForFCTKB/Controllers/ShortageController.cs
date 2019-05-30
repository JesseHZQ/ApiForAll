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
    public class ShortageController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpPost]
        public string UpdateShortage(SlotShortage shortage)
        {
            string update = "UPDATE KANBAN_SLOTSHORTAGE SET IsReceived = @IsReceived, ETA = @ETA WHERE ID = @ID";
            conn.Execute(update, shortage);
            return "ok";
        }
    }
}
