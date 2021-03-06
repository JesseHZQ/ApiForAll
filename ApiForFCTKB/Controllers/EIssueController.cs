﻿using System;
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
    public class EIssueController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["FCTKB"].ConnectionString);

        [HttpGet]
        public List<SlotEIssue> GetEIssueBySlot(string slot)
        {
            string sql = "SELECT * FROM KANBAN_SLOTEISSUE WHERE Slot = '" + slot + "'";
            return conn.Query<SlotEIssue>(sql).ToList();
        }

        [HttpPost]
        public int InsertEIssue(SlotEIssue eIssue)
        {
            eIssue.Date = DateTime.Now;
            string insert = "INSERT INTO KANBAN_SLOTEISSUE (Slot, Item, Date, Status) VALUES (@Slot, @Item, @Date, @Status)";
            return conn.Execute(insert, eIssue);
        }

        [HttpPost]
        public int UpdateEIssue(SlotEIssue eIssue)
        {
            eIssue.LastUpdatedDate = DateTime.Now;
            string update = "UPDATE KANBAN_SLOTEISSUE SET Item = @Item, Date = @Date, Status = @Status, LastUpdatedDate = @LastUpdatedDate WHERE ID = @ID";
            return conn.Execute(update, eIssue);
        }

    }
}
