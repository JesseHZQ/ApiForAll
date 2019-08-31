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
    public class UserController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);
        
        [HttpGet]
        public List<User> GetUserList()
        {
            string sql = "SELECT * FROM [CT_User] WHERE IsDel = 0 ORDER BY UserName";
            return conn.Query<User>(sql).ToList();
        }

        [HttpGet]
        public User GetUserById(string Id)
        {
            string sql = "SELECT * FROM [CT_User] WHERE UserId = " + Id;
            return conn.QueryFirstOrDefault<User>(sql);
        }

        [HttpPost]
        public User Login(User user)
        {
            string sql = "SELECT * FROM [CT_User] WHERE UserId = @UserId AND Password = @Password";
            return conn.QueryFirstOrDefault<User>(sql, user);
        }

        [HttpPost]
        public int AddUser(User user)
        {
            string sqlQuery = "SELECT * FROM [CT_User] WHERE UserId = @UserId";
            if (conn.QueryFirstOrDefault<User>(sqlQuery, user) == null)
            {
                string sql = "INSERT INTO [CT_User] (UserId, UserName, Password, Department, LevelNum, Shift, IsLeave, IsLock, Skills, Phone, Email, BackUpId, IsDel) VALUES (@UserId, @UserName, @Password, @Department, @LevelNum, @Shift, @IsLeave, @IsLock, @Skills, @Phone, @Email, @BackUpId, @IsDel)";
                return conn.Execute(sql, user);
            }
            return 0;
            
        }

        [HttpPost]
        public int DeleteUser(User user)
        {
            string sql = "UPDATE [CT_User] SET IsDel = 1 WHERE UserId = @UserId";
            return conn.Execute(sql, user);
        }

        [HttpPost]
        public int UpdateUser(User user)
        {
            string sql = "UPDATE [CT_User] SET UserName = @UserName, Password= @Password, Department = @Department, LevelNum = @LevelNum, Shift = @Shift, IsLeave = @IsLeave, IsLock = @IsLock, Skills = @Skills, Phone = @Phone, Email = @Email, BackUpId = @BackUpId WHERE UserId = @UserId";
            return conn.Execute(sql, user);
        }
    }
}
