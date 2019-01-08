using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForMinioneMiscShipping.Controllers
{
    public class TrainingController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp get()
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM [UF_Training].[dbo].[FCT_Master] a join [SC3_test].[dbo].[ReTraining] b on a.Topic = b.Topic");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public RespAll getDataList()
        {
            DataTable dt_User = new DataTable();
            DataTable dt_Type = new DataTable();
            DataTable dt_Detail = new DataTable();
            dt_User = SqlHelper.ExecuteDataTable("SELECT * FROM FCT_User");
            dt_Type = SqlHelper.ExecuteDataTable("SELECT * FROM FCT_Type");
            dt_Detail = SqlHelper.ExecuteDataTable("SELECT * FROM FCT_Detail");
            RespAll resp = new RespAll();
            resp.Code = 200;
            resp.Data_User = dt_User;
            resp.Data_Type = dt_Type;
            resp.Data_Detail = dt_Detail;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp InsertOrUpdate(string name, int stationId, int status)
        {
            Resp resp = new Resp();
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("SELECT * FROM FCT_Detail WHERE NAME = '" + name + "' AND TYPEID = " + stationId);
            if (dt.Rows.Count > 0)
            {
                int count = SqlHelper.ExecuteNonQuery("update FCT_Detail set status = " + status + ", MODIFIEDDATE = '" + DateTime.Now + "' WHERE NAME = '" + name + "' AND TYPEID = " + stationId);
                if (count > 0)
                {
                    resp.Code = 200;
                    resp.Message = "Success";
                    resp.Data = null;
                }
                else
                {
                    resp.Code = 10001;
                    resp.Message = "fail";
                    resp.Data = null;
                }
            }
            else
            {
                int count = SqlHelper.ExecuteNonQuery("insert into FCT_Detail (name, typeid, status, modifieddate) values('" + name + "', " + stationId + ", " + status + ",'" + DateTime.Now + "')");
                if (count > 0)
                {
                    resp.Code = 200;
                    resp.Message = "Success";
                    resp.Data = null;
                }
                else
                {
                    resp.Code = 10001;
                    resp.Message = "fail";
                    resp.Data = null;
                }
            }
            return resp;
        }
        public class Record
        {

        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }

        public class RespAll
        {
            public int Code { get; set; }
            public DataTable Data_User { get; set; }
            public DataTable Data_Type { get; set; }
            public DataTable Data_Detail { get; set; }
            public string Message { get; set; }
        }
    }
}
