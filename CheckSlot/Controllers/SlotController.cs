using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CheckSlot.DataContext;

namespace CheckSlot.Controllers
{
    public class SlotController : ApiController
    {
        SlotDemandDataContext dc = new SlotDemandDataContext();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sn"></param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp getData(Arr model)
        {
            Resp resp = new Resp();
            List<SlotDemandRequestDto> list = new List<SlotDemandRequestDto>();
            try
            {
                list = dc.SlotDemandRequestDto.Where(x => x.ItemNo == "TDN-425-240-00"
                || x.ItemNo == "TDN-640-128-00"
                || x.ItemNo == "TDN-361-749-03"
                || x.ItemNo == "TDN-429-291-01"
                || x.ItemNo == "TDN-361-749-04"
                || x.ItemNo == "TDN-361-749-05"
                || x.ItemNo == "TDN-361-749-08"
                ).ToList();
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
            return resp;
        }

    }

    public class Resp
    {
        public int Code { get; set; }
        public List<SlotDemandRequestDto> Data { get; set; }
        public string Message { get; set; }
    }

    public class Arr
    {
        public List<string> list { get; set; }
    }
}
