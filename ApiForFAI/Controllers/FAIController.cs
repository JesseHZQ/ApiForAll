using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForFAI.Controllers
{
    public class FAIController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp getWorkorder()
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select distinct Workorder from QESystem");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getIdByWorkorder(string workorder)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select systemId from QESystem where workorder = '" + workorder + "'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getContentById(string systemId)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select Station from QESystem where systemId = '" + systemId + "'");
            string tableName = "";
            if (dt.Rows[0]["Station"].ToString() == "SMT")
            {
                tableName = "QESMT";
            } else
            {
                tableName = "QEPTH";
            }
            DataTable dtContent = new DataTable();
            dtContent = SqlHelper.ExecuteDataTable("select * from " + tableName + " where systemId = '" + systemId + "' order by id");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dtContent;
            resp.Message = "查询成功";
            return resp;
        }
        
        [System.Web.Http.HttpPost]
        public Resp updateById(Model model)
        {
            string tableName = "";
            if (model.ListBy == "SMT")
            {
                tableName = "QESMT";
            }
            else
            {
                tableName = "QEPTH";
            }
            int index = SqlHelper.ExecuteNonQuery("update " + tableName + " set ModelName = '" + model.ModelName + "', P_F_Value_NA = '" + model.P_F_Value_NA + "', Result = '" + model.Result + "' where Num = " + model.Num);
            Resp resp = new Resp();
            if (index >= 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "执行成功";
            } else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "执行失败";
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getAllSystem(string date)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select * from QESystem where FORMAT(folderSetupDate,'yyyy-MM-dd') like '%" + date + "%'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updateSystemById(SystemInfo model)
        {
            int index = SqlHelper.ExecuteNonQuery("update QESystem set ModelName = '" + model.ModelName + "', Workorder = '" + model.Workorder + "', FAIType = '" + model.FAIType + "' where systemId = '" + model.systemId + "'");
            Resp resp = new Resp();
            if (index >= 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "执行成功";
            }
            else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "执行失败";
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getFailureSummary()
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select * from QEOpenIssue");
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

        public class Model
        {
            public int Num { get; set; }
            public string ListBy { get; set; }
            public string ModelName { get; set; }
            public string P_F_Value_NA { get; set; }
            public string Result { get; set; }
        }

        public class SystemInfo
        {
            public string ModelName { get; set; }
            public string Workorder { get; set; }
            public string FAIType { get; set; }
            public string systemId { get; set; }
        }
    }
}
