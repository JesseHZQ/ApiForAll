using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForRackManage.Controllers
{
    public class SystemController : ApiController
    {
        RackDataContext dc = new RackDataContext();
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
    }

    public class Resp
    {
        public int Code { get; set; }
        public List<Instrument> DataUF { get; set; }
        public List<J750_Board> Data750 { get; set; }
        public string Message { get; set; }
    }
}
