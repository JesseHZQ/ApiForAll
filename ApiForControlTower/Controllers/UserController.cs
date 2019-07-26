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
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User]";
            return conn.Query<User>(sql).ToList();
        }

        [HttpPost]
        public User Login(User user)
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User] WHERE UserID = @UserID AND Password = @Password";
            return conn.QueryFirstOrDefault<User>(sql, user);
        }

        [HttpGet]
        public List<User> GetUserById(string Id)
        {
            string sql = "SELECT * FROM [SlotKB].[dbo].[CT_User] WHERE UserID = " + Id;
            return conn.Query<User>(sql).ToList();
        }

        [HttpPost]
        public int UpdateUser(User user)
        {
            string sql = "UPDATE [SlotKB].[dbo].[CT_User] SET UserName = @UserName, Shift = @Shift, Station = @Station, Issue = @Issue, Phone = @Phone, Email = @Email WHERE Id = @Id";
            return conn.Execute(sql, user);
        }

        [HttpPost]
        public int AddUser(User user)
        {
            string sql = "INSERT INTO [SlotKB].[dbo].[CT_User] (UserId, UserName, Password, Shift, Phone, Email) VALUES (@UserId, @UserName, @Password, @Shift, @Phone, @Email)";
            return conn.Execute(sql, user);
        }

        [HttpPost]
        public int DeleteUser(User user)
        {
            string sql = "DELETE FROM [SlotKB].[dbo].[CT_User] WHERE Id = @Id";
            return conn.Execute(sql, user);
        }
    }
}
