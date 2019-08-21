using ApiForFCTKB.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ApiForFCTKB.Controllers
{
    public class DragonController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public List<Dragon> GetDragons()
        {
            string sql = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate IS NULL AND TYPE = 'D'";
            return conn.Query<Dragon>(sql).ToList();
        }

        [HttpPost]
        public int UpdateDragon(Dragon dragon)
        {
            string sql = "UPDATE KANBAN_SLOTPLAN SET Launch = @Launch, CoreBU = @CoreBU, PV = @PV, OI = @OI, TestBU = @TestBU, CSW = @CSW, QFAA = @QFAA, BU = @BU, Pack = @Pack WHERE ID = @ID";
            return conn.Execute(sql, dragon);
        }
    }
}
