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

namespace ApiForFCTKBNew.Controllers
{
    public class UserController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpPost]
        public UserInfo Login(UserInfo user)
        {
            string sql = "SELECT * FROM [CT_User] WHERE UserId = @UserId AND Password = @Password";
            return conn.QueryFirstOrDefault<UserInfo>(sql, user);
        }

        public class UserInfo
        {
            public int UserId { get; set; }
            public string UserName { get; set; }
            public string Password { get; set; }
            public string Department { get; set; }
            public int LevelNum { get; set; }
            public string Shift { get; set; }
            public int IsLeave { get; set; }
            public int IsLock { get; set; }
            public string Skills { get; set; }
            public string Phone { get; set; }
            public string Email { get; set; }
            public int BackUpId { get; set; }
            public int IsDel { get; set; }
        }
    }
}
