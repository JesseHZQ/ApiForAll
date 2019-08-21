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
    public class TaskController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        /// <summary>
        /// 获取所有User
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public List<TaskWork> GetTaskList()
        {
            string sql = "SELECT A.*, B.UserName FROM [CT_Task] A LEFT JOIN [CT_User] B ON A.OwnerId = B.UserId";
            return conn.Query<TaskWork>(sql).ToList();
        }

        [HttpPost]
        public int UpdateTask(TaskWork task)
        {
            string sql = "UPDATE [CT_Task] SET BayName = @BayName, SystemName= @SystemName, " +
                "FailureStation = @FailureStation, FailureType = @FailureType, FailureDesc = @FailureDesc, " +
                "OwnerId = @OwnerId, Status = @Status, RootCause = @RootCause, Solution = @Solution, " +
                "CreationTime = @CreationTime, CloseTime = @CloseTime, Creator = @Creator WHERE Id = @Id";
            return conn.Execute(sql, task);
        }

        [HttpPost]
        public int AddTask(TaskWork task)
        {
            string sql = "INSERT INTO [CT_Task] (BayName, SystemName, FailureStation, FailureType, FailureDesc, OwnerId, Status, RootCause, Solution, CreationTime, CloseTime, Creator) " +
                "VALUES (@BayName, @SystemName, @FailureStation, @FailureType, @FailureDesc, @OwnerId, @Status, @RootCause, @Solution, @CreationTime, @CloseTime, @Creator)";
            return conn.Execute(sql, task);
        }

        [HttpPost]
        public int DeleteTask(TaskWork task)
        {
            string sql = "DELETE FROM [CT_Task] WHERE Id = @Id";
            return conn.Execute(sql, task);
        }
    }
}
