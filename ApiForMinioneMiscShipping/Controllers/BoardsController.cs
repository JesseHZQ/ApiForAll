using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GROUP.Manage;

namespace ApiForMinioneMiscShipping.Controllers
{
    public class BoardsController : ApiController
    {
        BaseClass bc = new BaseClass();

        [System.Web.Http.HttpPost]
        public Resp get(Board board)
        {
            DataTable dt = new DataTable();
            if (board.sn == null)
            {
                dt = bc.ReadTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar]");
            } else
            {
                dt = bc.ReadTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN = '" + board.sn + "'");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        public class Board
        {
            public int start { get; set; }
            public int length { get; set; }
            public string sn { get; set; }

        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }
    }
}
