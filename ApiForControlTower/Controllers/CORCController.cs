using ApiForControlTower.Models;
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
            string sql = "select a.PD, b.* from KANBAN_SLOTPLAN a join KANBAN_SLOTCORC b on a.Slot = b.Slot order by b.InsertTime DESC";
            return connSlot.Query<CORC>(sql).ToList();
        }

        [HttpGet]
        public DateCORC GetDateCORC()
        {
            string sql = "select a.PD, b.* from KANBAN_SLOTPLAN a join KANBAN_SLOTCORC b on a.Slot = b.Slot";
            List<CORC> listCORC = connSlot.Query<CORC>(sql).ToList();

            // Day
            List<DayCORC> listDay = new List<DayCORC>();
            foreach (CORC corc in listCORC)
            {
                string date = corc.InsertTime.Year.ToString() + corc.InsertTime.Month.ToString() + corc.InsertTime.Day.ToString();
                DayCORC dc = listDay.Where(x => x.Date == date).SingleOrDefault();
                if (dc == null)
                {
                    List<CORC> clist = new List<CORC>();
                    clist.Add(corc);
                    listDay.Add(new DayCORC
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
            listDay = listDay.OrderBy(item => item.Date).ToList();

            // Week
            List<WeekCORC> listWeek = new List<WeekCORC>();
            foreach (CORC corc in listCORC)
            {
                string week = Tool.WeekOfYear(corc.InsertTime).ToString();
                string date = corc.InsertTime.Year.ToString() + (week.Length == 2 ? week : "0" + week);
                WeekCORC dc = listWeek.Where(x => x.Date == date).SingleOrDefault();
                if (dc == null)
                {
                    List<CORC> clist = new List<CORC>();
                    clist.Add(corc);
                    listWeek.Add(new WeekCORC
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
            listWeek = listWeek.OrderBy(item => item.Date).ToList();

            // Month
            List<MonthCORC> listMonth = new List<MonthCORC>();
            foreach (CORC corc in listCORC)
            {
                string date = corc.InsertTime.Year.ToString() + corc.InsertTime.Month.ToString();
                MonthCORC dc = listMonth.Where(x => x.Date == date).SingleOrDefault();
                if (dc == null)
                {
                    List<CORC> clist = new List<CORC>();
                    clist.Add(corc);
                    listMonth.Add(new MonthCORC
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
            listMonth = listMonth.OrderBy(item => item.Date).ToList();
            return new DateCORC {
                DayCorc = listDay,
                WeekCorc = listWeek,
                MonthCorc = listMonth
            };
        }

        public class DayCORC
        {
            public string Date { get; set; }
            public List<CORC> Corc { get; set; }
            public int CorcAmount { get; set; }
        }

        public class WeekCORC
        {
            public string Date { get; set; }
            public List<CORC> Corc { get; set; }
            public int CorcAmount { get; set; }
        }

        public class MonthCORC
        {
            public string Date { get; set; }
            public List<CORC> Corc { get; set; }
            public int CorcAmount { get; set; }
        }

        public class DateCORC
        {
            public List<DayCORC> DayCorc { get; set; }
            public List<WeekCORC> WeekCorc { get; set; }
            public List<MonthCORC> MonthCorc { get; set; }
        }
    }
}
