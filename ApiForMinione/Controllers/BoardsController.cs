using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForMinione.Controllers
{
    public class BoardsController : ApiController
    {
        [System.Web.Http.HttpPost]
        public Resp get(Board board)
        {
            //SELECT* FROM[172.21.194.214].[PCBA].[dbo].[Failuar] where Location = 'LeanProcess' and Date between '2018-12-10' and '2018-12-17' and Station<> 'NULL' order by Date
            DataTable dt = new DataTable();
            if (board.sn == null)
            {
                dt = SqlHelper.ExecuteDataTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar]");
            }
            else
            {
                dt = SqlHelper.ExecuteDataTable("SELECT TOP " + board.length + " * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN = '" + board.sn + "'");
            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        public Resp get(string sn)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN = '" + sn + "'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        public Resp getSnList(string sn)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT distinct SN FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN like '%" + sn + "%'");
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
