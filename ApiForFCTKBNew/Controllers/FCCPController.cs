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
    public class FCCPController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public string UpdateFCCP()
        {
            try
            {
                string sqlFCCP = "select * from openquery(FCCP, 'SELECT * FROM FCCP011.V_ITEM_PRODUCTION_011_SHIPMENT')";
                List<FCCP> listFCCP = conn.Query<FCCP>(sqlFCCP).ToList();

                string sqlSlotPlan = "SELECT * FROM KANBAN_SLOTPLAN WHERE ISNULL(ShippingDate, 0) = '0' ORDER BY MRP";
                List<SlotPlan> listSlotPlan = conn.Query<SlotPlan>(sqlSlotPlan).ToList().Where(x => Tool.GetValidate(x.MRP) == true).ToList();

                string sqlType = "select * from KANBAN_FCCP_PN2TYPE";
                List<PN2TYPE> listType = conn.Query<PN2TYPE>(sqlType).ToList();

                foreach (SlotPlan sp in listSlotPlan)
                {
                    FCCP fccp = listFCCP.Where(x => x.CODE.Contains(sp.Slot)).FirstOrDefault();
                    if (fccp == null)
                    {
                        sp.FCCPStatus = 0;
                        sp.TypeStatus = 0;
                        sp.IBFStatus = null;
                    }
                    else
                    {
                        sp.FCCPStatus = 1;
                        sp.IBFStatus = fccp.OVER_FLAG;
                        PN2TYPE type = listType.Where(x => x.SlotType == sp.Model).FirstOrDefault();
                        if (type != null)
                        {
                            if (fccp.REQ_MODEL != null && fccp.REQ_MODEL.Contains(type.PN))
                            {
                                sp.TypeStatus = 1;
                            }
                            else
                            {
                                sp.TypeStatus = 0;
                            }
                        }
                        else
                        {
                            sp.TypeStatus = 0;
                        }
                    }
                }

                string sql = "UPDATE KANBAN_SLOTPLAN SET FCCPStatus = @FCCPStatus, TypeStatus = @TypeStatus, IBFStatus = @IBFStatus WHERE Slot = @Slot";
                conn.Execute(sql, listSlotPlan);
                return "FCCP OK!";
            }
            catch (Exception ex)
            {
                return "FCCP Failed! " + ex.Message;
            }
        }

        private class FCCP
        {
            public string CODE { get; set; }
            public string REQ_MODEL { get; set; }
            public string OVER_FLAG { get; set; }
        }

        private class PN2TYPE
        {
            public int Id { get; set; }
            public string SlotType { get; set; }
            public string PN { get; set; }
        }
    }
}
