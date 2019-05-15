using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForMinioneMiscShipping.Controllers
{
    public class DragonController : ApiController
    {
        [System.Web.Http.HttpPost]
        public Resp UploadDragonData(SystemList model)
        {
            foreach (var item in model.list)
            {
                DataTable dt = new DataTable();
                dt = sqlkb.ExecuteDataTable("select * from SummaryHDragon where SystemModel = '" + item.Slot + "'");
                if (dt.Rows.Count == 0)
                {
                    sqlkb.ExecuteNonQuery("Insert into SummaryHDragon (SystemSlot, SystemModel, Customer, PO, SO, ShipWeek, ShipDate, Lock, IsMore) values ('" + item.Slot + "', '" + item.Model + "','" + item.Customer + "','" + item.PO + "', '" + item.SO + "','" + item.MRP + "', '" + item.PD + "', 0, 'True')");
                }
                else
                {
                    sqlkb.ExecuteNonQuery("Update SummaryHDragon set PO = '" + item.PO + "', SO = '" + item.SO + "', ShipWeek = '" + item.MRP + "', ShipDate = '" + item.PD + "' where SystemSlot = " + item.Slot);
                }
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Message = "Success";
            resp.Data = null;
            return resp;
        }

        public class SystemInfo
        {
            public string Slot { get; set; }
            public string Model { get; set; }
            public string Customer { get; set; }
            public string PO { get; set; }
            public string PD { get; set; }
            public string MRP { get; set; }
            public string SO { get; set; }
        }

        public class SystemList
        {
            public List<SystemInfo> list { get; set; }
        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }
    }
}
