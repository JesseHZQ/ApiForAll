using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;

namespace ApiForRackManage.Controllers
{
    public class RackController : ApiController
    {
        RackDataContext dc = new RackDataContext();

        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        [System.Web.Http.HttpGet]
        public Resp getList(string PN)
        {
            Resp resp = new Resp();
            if (PN == "0")
            {
                List<Rack> list = new List<Rack>();
                try
                {
                    list = dc.Rack.ToList();
                    if (list != null)
                    {
                        resp.Code = 200;
                        resp.Data = list;
                        resp.Message = "success";
                    }
                    else
                    {
                        resp.Code = 10001;
                        resp.Data = null;
                        resp.Message = "No Data";
                    }
                }
                catch (Exception ex)
                {
                    resp.Code = 10001;
                    resp.Data = null;
                    resp.Message = ex.Message;
                }
            }
            else
            {
                List<Rack> list = new List<Rack>();
                try
                {
                    list = dc.Rack.Where(x => x.PN == PN).ToList();
                    if (list.Count != 0)
                    {
                        resp.Code = 200;
                        resp.Data = list;
                        resp.Message = "success";
                    }
                    else
                    {
                        resp.Code = 10001;
                        resp.Data = null;
                        resp.Message = "No Data";
                    }
                }
                catch (Exception ex)
                {
                    resp.Code = 10001;
                    resp.Data = null;
                    resp.Message = ex.Message;
                }
            }
            return resp;
        }


        [HttpGet]
        public List<Rac> getRaList()
        {
            Resp resp = new Resp();
            string sql = "select * from Supermarket_Rack";
            return conn.Query<Rac>(sql).ToList();
        }
        public class Rac
        {
            public int RackId { get; set; }
            public string RackName { get; set; }
            public int RackSize { get; set; }
            public int TypeId { get; set; }
            public string PN { get; set; }
            public string PNList { get; set; }
            public string SNList { get; set; }
            public string SlotList { get; set; }
            public string DateList { get; set; }
        }

        [System.Web.Http.HttpPost]
        public Resp Update(Rack model)
        {
            Rack item = new Rack();
            Resp resp = new Resp();
            item = dc.Rack.FirstOrDefault(x => x.ID == model.ID);
            item.PN = model.PN;
            item.ActualQTY = model.ActualQTY;
            item.Size = model.Size;
            item.SNView = model.SNView;
            item.SlotView = model.SlotView;
            item.TimeView = model.TimeView;
            item.Command = model.Command;
            dc.SubmitChanges();
            resp.Code = 200;
            resp.Data = null;
            resp.Message = "更新成功";
            return resp;
        }

        public class Resp
        {
            public int Code { get; set; }
            public List<Rack> Data { get; set; }
            public string Message { get; set; }
        }
    }
}
