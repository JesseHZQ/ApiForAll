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
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["suznt004"].ConnectionString);

        /// <summary>
        /// 获取所有User
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public List<User> GetUserList()
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User] ORDER BY UserName";
            return conn.Query<User>(sql).ToList();
        }

        [HttpPost]
        public User Login(User user)
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User] WHERE UserId = @UserId AND Password = @Password";
            return conn.QueryFirstOrDefault<User>(sql, user);
        }

        [HttpGet]
        public User GetUserById(string Id)
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User] WHERE UserId = " + Id;
            return conn.QueryFirstOrDefault<User>(sql);
        }

        [HttpPost]
        public int UpdateUser(User user)
        {
            string sql = "UPDATE [SlotKB].[dbo].[CT_User] SET UserName = @UserName, Password= @Password, LevelNum = @LevelNum, Shift = @Shift, Skills = @Skills, Phone = @Phone, Email = @Email, BackUpId = @BackUpId WHERE UserId = @UserId";
            return conn.Execute(sql, user);
        }

        [HttpPost]
        public int AddUser(User user)
        {
            string sqlQuery = "SELECT * FROM [SlotKB].[dbo].[CT_User] WHERE UserId = @UserId";
            User u = conn.QueryFirstOrDefault<User>(sqlQuery, user);
            if (u == null)
            {
                string sql = "INSERT INTO [SlotKB].[dbo].[CT_User] (UserId, UserName, Password, LevelNum, Shift, Skills, Phone, Email, BackUpId) VALUES (@UserId, @UserName, @Password, @LevelNum, @Shift, @Skills, @Phone, @Email, @BackUpId)";
                return conn.Execute(sql, user);
            } else
            {
                return 0;
            }
            
        }

        [HttpPost]
        public int DeleteUser(User user)
        {
            string sql = "DELETE FROM [SlotKB].[dbo].[CT_User] WHERE UserId = @UserId";
            return conn.Execute(sql, user);
        }
    }
}
