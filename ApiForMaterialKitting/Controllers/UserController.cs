using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForMaterialKitting.Controllers
{
    public class UserController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp getUserName(string emp)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select Name from MaterialKittingUser where Emp = '" + emp + "'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        public class Resp
        {
            public int Code { get; set; }
            public DataTable Data { get; set; }
            public string Message { get; set; }
        }

    }
}
