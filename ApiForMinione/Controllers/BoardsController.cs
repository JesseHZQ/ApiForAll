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
        public Resp getInfoBySN(string sn)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Failuar] a JOIN [172.21.194.214].[PCBA].[dbo].[Boards] b ON a.SN = b.SN WHERE a.SN = '" + sn + "' ORDER BY a.Date");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getInfoByDate(long begin, long end)
        {
            DateTime dt1 = ConvertTime(begin);
            DateTime dt2 = ConvertTime(end);
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Failuar] a JOIN [172.21.194.214].[PCBA].[dbo].[Boards] b ON a.SN = b.SN WHERE ((a.Location = 'LeanProcess' and a.NextLocation = 'LeanRack') or (a.Location = 'LeanProcess' and a.NextLocation = 'Verifying') or (a.Location = 'PreBurnIn') or (a.Location = 'Misc' and a.NextLocation = 'MC') or (a.Location = 'Misc' and a.NextLocation = 'Verifying')) and (Date between '" + dt1 + "' and '" + dt2 + "') ORDER BY a.SN");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getSlotInfoByDate(long begin, long end)
        {
            DateTime dt1 = ConvertTime(begin);
            DateTime dt2 = ConvertTime(end);
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Failuar] a JOIN [172.21.194.214].[PCBA].[dbo].[Boards] b ON a.SN = b.SN WHERE ((a.Location = 'Shipping' and a.NextLocation = 'Shipping') or (a.Location = 'tester' and a.NextLocation = 'Verifying ')) and (Date between '" + dt1 + "' and '" + dt2 + "') ORDER BY a.SN");
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
            dt = SqlHelper.ExecuteDataTable("SELECT distinct top 1000 SN FROM[172.21.194.214].[PCBA].[dbo].[Failuar] WHERE SN like '%" + sn + "%'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getSystemList()
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT distinct TesterID FROM[172.21.194.214].[PCBA].[dbo].[Failuar] order by TesterID");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp getInfoBySystem(SystemListModel model)
        {
            string str = "(";
            foreach (var item in model.systemList)
            {
                str += "'" + item + "',";
            }
            str = str.Substring(0, str.Length - 1);
            str += ")";
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM[172.21.194.214].[PCBA].[dbo].[Failuar] a JOIN [172.21.194.214].[PCBA].[dbo].[Boards] b ON a.SN = b.SN where a.TesterID in " + str);
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        public DateTime ConvertTime(long milliTime)
        {
            long timeTricks = new DateTime(1970, 1, 1).Ticks + milliTime * 10000 + TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).Hours * 3600 * (long)10000000;
            return new DateTime(timeTricks);
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

        public class SystemListModel
        {
            public List<string> systemList { get; set; }
        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }
    }
}
