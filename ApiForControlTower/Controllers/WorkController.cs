using ApiForControlTower.Models;
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


namespace ApiForControlTower.Controllers
{
    public class WorkController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public List<Work> GetWorkList()
        {
            string sql = "SELECT * FROM CT_Work";
            return conn.Query<Work>(sql).ToList();
        }

        [HttpGet]
        public User GetWorkByOwnerId(string Id)
        {
            string sql = "SELECT * FROM CT_Work WHERE OwnerId = " + Id;
            return conn.QueryFirstOrDefault<User>(sql);
        }

        [HttpPost]
        public int AddWork(Work work)
        {
            string sql = "INSERT INTO CT_Work (BayName, SystemName, Station, Date, OwnerId) VALUES (@BayName, @SystemName, @Station, @Date, @OwnerId)";
            return conn.Execute(sql, work);
        }
    }
}