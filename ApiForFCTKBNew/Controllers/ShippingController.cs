using ApiForFCTKBNew.Models;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForFCTKBNew.Controllers
{
    public class ShippingController : ApiController
    {
        public IDbConnection connKB = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        public IDbConnection connTER = new SqlConnection(ConfigurationManager.ConnectionStrings["TER"].ConnectionString);

        public string UpdateShippingStatus()
        {
            try
            {
                // 获取所有未出货系统
                string sql = "SELECT * FROM KANBAN_SLOTPLAN WHERE ShippingDate IS NULL";
                List<SlotPlan> slotPlans = connKB.Query<SlotPlan>(sql).ToList();

                // 获取所有出货信息
                string sqlPO = "select [PO NO.] as PO, Status from vFFPOStatus";
                List<POStatus> POs = connTER.Query<POStatus>(sqlPO).ToList();

                foreach (SlotPlan sp in slotPlans)
                {
                    POStatus ps = POs.Where(x => x.PO == sp.PO).FirstOrDefault();
                    if (ps != null && ps.PO == "Shipped")
                    {
                        sp.ShippingDate = Tool.GetFormatDateWithDay(DateTime.Now); // 2020年第五周星期四 ==> 202005.4
                    }
                }

                string updateSql = "update KANBAN_SLOTPLAN set ShippingDate = @ShippingDate where ID = @ID";
                connKB.Execute(updateSql, slotPlans);
                return "Shipping OK!";
            }
            catch (Exception ex)
            {
                return "Shipping Failed!" + ex.Message;
            }
            
        }

        private class POStatus
        {
            public string PO { get; set; }
            public string Status { get; set; }
        }
    }
}
