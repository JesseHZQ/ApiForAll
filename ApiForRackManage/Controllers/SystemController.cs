using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ApiForRackManage.Controllers
{
    public class SystemController : ApiController
    {
        RackDataContext dc = new RackDataContext();
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);
        public IDbConnection connPNINOUT = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        [System.Web.Http.HttpGet]
        public Resp getSystemList()
        {
            Resp resp = new Resp();
            List<Instrument> listUF = new List<Instrument>();
            List<J750_Board> list750 = new List<J750_Board>();
            listUF = dc.Instrument.ToList();
            list750 = dc.J750_Board.ToList();
            resp.Code = 200;
            resp.DataUF = listUF;
            resp.Data750 = list750;
            resp.Message = "success";
            return resp;
        }

        [HttpPost]
        public int UpdateSystem(SysList list)
        {
            string sql = "UPDATE Instrument SET QTY = @total, BI = @BI where PN = @PN";
            return connPNINOUT.Execute(sql, list.list);
        }

        [HttpPost]
        public int Update750System(SysList list)
        {
            string sql = "UPDATE J750_Board SET QTY = @total, BI = @BI where PN = @PN";
            return connPNINOUT.Execute(sql, list.list);
        }

        [System.Web.Http.HttpGet]
        public RespBI getBurnIn()
        {
            string sql = "SELECT PN, SN FROM [192.168.163.1].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            RespBI resp = new RespBI();
            List<BI> list = connPNINOUT.Query<BI>(sql).ToList();
            resp.List = list;
            resp.Message = "success";
            return resp;
        }
    }

    public class Sys
    {
        public string PN { get; set; }
        public int total { get; set; }
        public int BI { get; set; }
    }

    public class SysList
    {
        public List<Sys> list { get; set; }
    }

    public class Resp
    {
        public int Code { get; set; }
        public List<Instrument> DataUF { get; set; }
        public List<J750_Board> Data750 { get; set; }
        public string Message { get; set; }
    }

    public class RespBI
    {
        public int Code { get; set; }
        public List<BI> List { get; set; }
        public string Message { get; set; }
    }

    public class BI
    {
        public string PN { get; set; }
        public string SN { get; set; }
    }
}
