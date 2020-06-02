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

namespace ApiForFCTKBNew.Controllers
{
    public class DragonController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public List<Dragon> GetDragons()
        {
            string sql = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate IS NULL ORDER BY Slot";
            return conn.Query<Dragon>(sql).ToList();
        }

        [HttpPost]
        public int UpdateDragon(Dragon dragon)
        {
            string sql = "UPDATE KANBAN_SLOTPLAN SET Launch = @Launch, CoreBU = @CoreBU, PV = @PV, OI = @OI, TestBU = @TestBU, CSW = @CSW, QFAA = @QFAA, BU = @BU, Pack = @Pack WHERE ID = @ID";
            return conn.Execute(sql, dragon);
        }

        public class Dragon
        {
            public int ID { get; set; }
            public string Slot { get; set; }
            public string Launch { get; set; }
            public string CoreBU { get; set; }
            public string PV { get; set; }
            public string OI { get; set; }
            public string TestBU { get; set; }
            public string CSW { get; set; }
            public string QFAA { get; set; }
            public string BU { get; set; }
            public string Pack { get; set; }
        }
    }
}
