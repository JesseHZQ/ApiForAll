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
    public class FailureController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);
        
        [HttpGet]
        public List<Failure> GetFailureList()
        {
            string sql = "SELECT * FROM [CT_FailureList]";
            return conn.Query<Failure>(sql).ToList();
        }

        [HttpGet]
        public List<Failure> GetFailureStationList()
        {
            string sql = "SELECT DISTINCT FailureStation FROM [CT_FailureList]";
            return conn.Query<Failure>(sql).ToList();
        }

        [HttpGet]
        public List<Failure> GetFailureTypeList()
        {
            string sql = "SELECT DISTINCT FailureStation, FailureType FROM [CT_FailureList]";
            return conn.Query<Failure>(sql).ToList();
        }

        [HttpPost]
        public int UpdateFailure(Failure failure)
        {
            string sql = "UPDATE [CT_FailureList] SET FailureStation = @FailureStation, FailureType= @FailureType, FailureDesc = @FailureDesc WHERE Id = @Id";
            return conn.Execute(sql, failure);
        }

        [HttpPost]
        public int AddFailure(Failure failure)
        {
            string sql = "INSERT INTO [CT_FailureList] (FailureStation, FailureType, FailureDesc) VALUES (@FailureStation, @FailureType, @FailureDesc)";
            return conn.Execute(sql, failure);
        }

        [HttpPost]
        public int DeleteFailure(Failure failure)
        {
            string sql = "DELETE FROM [CT_FailureList] WHERE Id = @Id";
            return conn.Execute(sql, failure);
        }
    }
}
