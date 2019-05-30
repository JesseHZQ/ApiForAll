﻿using ApiForRackManage.DataContext;
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

        [System.Web.Http.HttpGet]
        public RespBI getBurnIn()
        {
            string sql = "SELECT PN, SN FROM[172.21.194.214].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            RespBI resp = new RespBI();
            List<BI> list = conn.Query<BI>(sql).ToList();
            resp.List = list;
            resp.Message = "success";
            return resp;
        }
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
