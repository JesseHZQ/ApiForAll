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
        public Resp getData(Board board)
        {
            //SELECT* FROM[172.21.194.214].[PCBA].[dbo].[Failuar] where Location = 'LeanProcess' and Date between '2018-12-10' and '2018-12-17' and Station<> 'NULL' order by Date
            DataTable dt = new DataTable();
            if (board.SN == null)
            {
                dt = SqlHelper.ExecuteDataTable("SELECT * FROM[172.21.194.214].[PCBA].[dbo].[Failuar]");
            }
            else
            {
                //dt = SqlHelper.ExecuteDataTable("SELECT * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN = '" + board.SN + "'");
                dt = SqlHelper.ExecuteDataTable("SELECT * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] where Location = 'LeanProcess' and Date between '2018-12-10' and '2018-12-17' and Station <> 'NULL' order by Date");

            }
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
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

        [System.Web.Http.HttpGet]
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
            public string SN { get; set; }
            public string Location { get; set; }
            public string NextLocation { get; set; }
            public string TesterID { get; set; }
            public string Station { get; set; }
            public DateTime Date { get; set; }
            public string Owner { get; set; }

        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }
    }
}
