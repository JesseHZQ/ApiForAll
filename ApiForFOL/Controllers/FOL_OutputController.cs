using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;


namespace ApiForFOL.Controllers
{
    public class FOL_OutputController : ApiController
    {
        [System.Web.Http.HttpGet]
        public RespAll getAllDetail(string version)
        {
            RespAll resp = new RespAll();
            resp.PriorForcast = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_PriorFsct] where version='" + version + "'");
            resp.CurrentForcast = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_CurFcst] where version='" + version + "'");
            resp.Actual = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_Actuals] where version='" + version + "'");
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getDetail(string version)
        {
            Resp resp = new Resp();
            try
            {
                resp.Data = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_PriorFsct] where version='" + version + "'");
                resp.Code = "200";
                resp.Message = "success";
            }
            catch (Exception ex)
            {
                resp.Code = "10001";
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }

        public Resp getActuals(string version)
        {
            Resp resp = new Resp();
            try
            {
                resp.Data = SqlHelper.ExecuteDataTable("select * from[dbo].[FOL_OutputByProject_Actuals] where version = '" + version + "'");
                resp.Code = "200";
                resp.Message = "success";
            }
            catch (Exception ex)
            {
                resp.Code = "10000";
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }
    }
}

public class Resp
{
    public string Code { get; set; }
    public DataTable Data { get; set; }
    public string Message { get; set; }
}

public class RespAll
{
    public string Code { get; set; }
    public DataTable PriorForcast { get; set; }
    public DataTable CurrentForcast { get; set; }
    public DataTable Actual { get; set; }
    public string Message { get; set; }
}