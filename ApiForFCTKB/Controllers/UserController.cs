using ApiForFCTKB.Models;
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

namespace ApiForFCTKB.Controllers
{
    public class UserController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);
        
        [HttpPost]
        public User Login(User user)
        {
            string sql = "SELECT * FROM [CT_User] WHERE UserId = @UserId AND Password = @Password";
            return conn.QueryFirstOrDefault<User>(sql, user);
        }
    }
}
