using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForCOQ.Controllers
{
    public class COQListController : ApiController
    {
        [System.Web.Http.HttpPost]
        public Resp getListByTime(QueryTime model)
        {
            DataTable dt = new DataTable();
            dt = SqlHelperList.ExecuteDataTable("select SN, PartNumber,ProductionOrderNumber,EnterTime,StationTypeID,StationType from SF_FF_History_Result_O where EnterTime between '" + model.beginTime +" ' and '" + model.finishTime + "' order by EnterTime");
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

        public class QueryTime
        {
            public string beginTime { get; set; }
            public string finishTime { get; set; }
        }
    }
}
