using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForFOL.Controllers
{
    public class FOL_InputController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp getTypedList(string version, string type)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_1_1] where version= '" + version + " ' and type like '%" + type + "%'");
            Resp resp = new Resp();
            resp.Code = "200";
            resp.Data = dt;
            resp.Message = "查询成功!";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getItemById(int id)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_1_1] where id = " + id);
            Resp resp = new Resp();
            resp.Code = "200";
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updateItemById(Input model)
        {
            int index;
            index = SqlHelper.ExecuteNonQuery("update FOL_Input_1_1 set " +
            "Period1 = " + model.Period1 + ", " +
            "Period2 = " + model.Period2 + ", " +
            "Period3 = " + model.Period3 + ", " +
            "Period4 = " + model.Period4 + ", " +
            "Period5 = " + model.Period5 + ", " +
            "Period6 = " + model.Period6 + ", " +
            "Period7 = " + model.Period7 + ", " +
            "Period8 = " + model.Period8 + ", " +
            "Period9 = " + model.Period9 + ", " +
            "Period10 = " + model.Period10 + ", " +
            "Period11 = " + model.Period11 + ", " +
            "Period12 = " + model.Period12 + ", " +
            "Period13 = " + model.Period13 + ", " +
            "Period14 = " + model.Period14 + ", " +
            "Period15 = " + model.Period15 + ", " +
            "Period16 = " + model.Period16 +
            "where id = " + model.Id);
            Resp resp = new Resp();
            if (index != 0)
            {
                resp.Code = "200";
                resp.Data = null;
                resp.Message = "修改成功";
            } else
            {
                resp.Code = "10001";
                resp.Data = null;
                resp.Message = "修改失败";
            }
            
            return resp;
        }
    }

    public class Resp
    {
        public string Code { get; set; }
        public DataTable Data { get; set; }
        public string Message { get; set; }
    }

    public class Input
    {
        public int Id { get; set; }
        public float Period1 { get; set; }
        public float Period2 { get; set; }
        public float Period3 { get; set; }
        public float Period4 { get; set; }
        public float Period5 { get; set; }
        public float Period6 { get; set; }
        public float Period7 { get; set; }
        public float Period8 { get; set; }
        public float Period9 { get; set; }
        public float Period10 { get; set; }
        public float Period11 { get; set; }
        public float Period12 { get; set; }
        public float Period13 { get; set; }
        public float Period14 { get; set; }
        public float Period15 { get; set; }
        public float Period16 { get; set; }
    }
}
