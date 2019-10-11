using ApiForControlTower.Models;
using ApiForWechat.Controllers;
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
    public class WorkAssignController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public int AutoAssignMorning()
        {
            return AutoAssign("D");
        }

        [HttpGet]
        public int AutoAssignAfternoon()
        {
            return AutoAssign("S");
        }

        public int AutoAssign(string shift)
        {
            List<WorkAssign> list = new List<WorkAssign>();

            List<LED> LEDList = conn.Query<LED>("SELECT * FROM CT_LEDMaster WHERE Station <> ''").ToList();

            List<UserAssign> UserList = conn.Query<UserAssign>("SELECT * FROM CT_User WHERE Shift = '" + shift + "' AND IsLeave = 0 AND IsLock = 0 ").ToList();

            foreach (LED led in LEDList.Where(x => x.Station == "QFAA"))
            {
                foreach (UserAssign user in UserList)
                {
                    if (user.Skills.Contains("QFAA") && !user.IsAssignedQFAA)
                    {
                        user.IsAssignedQFAA = true;
                        user.WorkAmount++;
                        WorkAssign wa = new WorkAssign();
                        wa.Id = led.Id;
                        wa.BayName = led.BayName;
                        wa.SystemName = led.SystemName;
                        wa.Station = led.Station;
                        wa.Date = DateTime.Now;
                        wa.OwnerId = user.UserId;
                        list.Add(wa);
                        break;
                    }
                }
            }

            foreach (LED led in LEDList.Where(x => x.Station != "QFAA"))
            {
                List<UserAssign> userAssigns = UserList.Where(x => !x.IsAssignedQFAA).ToList();
                userAssigns.Sort((a, b) => a.WorkAmount - b.WorkAmount);
                foreach (UserAssign user in userAssigns)
                {
                    if (user.Skills.Contains(led.Station))
                    {
                        user.WorkAmount++;
                        WorkAssign wa = new WorkAssign();
                        wa.Id = led.Id;
                        wa.BayName = led.BayName;
                        wa.SystemName = led.SystemName;
                        wa.Station = led.Station;
                        wa.Date = DateTime.Now;
                        wa.OwnerId = user.UserId;
                        list.Add(wa);
                        break;
                    }
                }
            }

            string sqlRemoveWork = "update CT_LEDMaster set OwnerId = null where Station <> ''";
            conn.Execute(sqlRemoveWork, list);
            string sqlAddWork = "insert into CT_Work (BayName, SystemName, Station, Date, OwnerId) values (@BayName, @SystemName, @Station, @Date, @OwnerId)";
            string sqlAssignWork = "update CT_LEDMaster set OwnerId = @OwnerId where Id = @Id";
            conn.Execute(sqlAssignWork, list);
            conn.Execute(sqlAddWork, list);
            new SendAssignWorkMessageController().SendAssignWorkMessage();
            return 1;
        }


        public class WorkAssign
        {
            public int Id { get; set; }
            public string BayName { get; set; }
            public string SystemName { get; set; }
            public string Station { get; set; }
            public DateTime Date { get; set; }
            public int OwnerId { get; set; }
        }

        public class UserAssign : User
        {
            public int WorkAmount { get; set; }
            public bool IsAssignedQFAA { get; set; }
        }
    }
}
