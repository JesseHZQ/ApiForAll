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
    public class StationController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public List<Station> GetStationList()
        {
            string sql = "SELECT * FROM [CT_StationList]";
            return conn.Query<Station>(sql).ToList();
        }

        [HttpPost]
        public int UpdateStation(Station station)
        {
            string sql = "UPDATE [CT_StationList] SET Department = @Department, StationName= @StationName WHERE Id = @Id";
            return conn.Execute(sql, station);
        }

        [HttpPost]
        public int AddStation(Station station)
        {
            string sql = "INSERT INTO [CT_StationList] (Department, StationName) VALUES (@Department, @StationName)";
            return conn.Execute(sql, station);
        }

        [HttpPost]
        public int DeleteStation(Station station)
        {
            string sql = "DELETE FROM [CT_StationList] WHERE Id = @Id";
            return conn.Execute(sql, station);
        }
    }
}
