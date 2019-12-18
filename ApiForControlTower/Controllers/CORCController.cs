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

namespace ApiForControlTower.Controllers
{
    public class CORCController : ApiController
    {
        public IDbConnection connSlot = new SqlConnection(ConfigurationManager.ConnectionStrings["SlotKB"].ConnectionString);
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);

        [HttpGet]
        public List<CORC> GetCORC()
        {
            string sql = "select a.PD, b.* from KANBAN_SLOTPLAN a join KANBAN_SLOTCORC b on a.Slot = b.Slot";
            return connSlot.Query<CORC>(sql).ToList();
        }

        [HttpGet]
        public List<DateCORC> GetDateCORC()
        {
            string sql = "select a.PD, b.* from KANBAN_SLOTPLAN a join KANBAN_SLOTCORC b on a.Slot = b.Slot";
            List<CORC> listCORC = connSlot.Query<CORC>(sql).ToList();
            List<DateCORC> list = new List<DateCORC>();
            foreach (CORC corc in listCORC)
            {
                string date = corc.InsertTime.Year.ToString() + corc.InsertTime.Month.ToString() + corc.InsertTime.Day.ToString();
                DateCORC dc = list.Where(x => x.Date == date).SingleOrDefault();
                if (dc == null)
                {
                    List<CORC> clist = new List<CORC>();
                    clist.Add(corc);
                    list.Add(new DateCORC
                    {
                        Date = date,
                        Corc = clist,
                        CorcAmount = 1
                    });
                }
                else
                {
                    dc.Corc.Add(corc);
                    dc.CorcAmount++;
                }
            }
            list = list.OrderBy(item => item.Date).ToList();
            return list;
        }

        public class DateCORC
        {
            public string Date { get; set; }
            public List<CORC> Corc { get; set; }
            public int CorcAmount { get; set; }
        }
    }
}
