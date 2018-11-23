using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForRackManage.Controllers
{
    public class RackController : ApiController
    {
        RackDataContext dc = new RackDataContext();

        /// <summary>
        /// 获取首页报告 date传0时,获取所有记录  传具体年月时,获取对应的记录
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 添加或者修改VOP记录
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp AddOrUpdateVOP(Rack model)
        {
            Rack item = new Rack();
            Resp resp = new Resp();
            item = dc.Rack.FirstOrDefault(x => x.ID == model.ID);
            item.PN = model.PN;
            item.Size = model.Size;
            item.SNView = model.SNView;
            item.SlotView = model.SlotView;
            item.TimeView = model.TimeView;
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
