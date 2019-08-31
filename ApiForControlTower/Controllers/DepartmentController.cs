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
    public class DepartmentController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public List<Department> GetUserList()
        {
            string sql = "SELECT * FROM [CT_Department]";
            return conn.Query<Department>(sql).ToList();
        }

    }
}
