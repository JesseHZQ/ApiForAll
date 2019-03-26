using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForMaterialKitting.Controllers
{
    public class RackController : ApiController
    {
        [System.Web.Http.HttpGet]
        public Resp inRack(int rackId, int KittingId, string RackName)
        {
            int count = SqlHelper.ExecuteNonQuery("update MaterialKittingRack set kittingid = '" + KittingId + "' where id = '" + rackId + "'");
            int ct = SqlHelper.ExecuteNonQuery("update MaterialKitting set RackName = '" + RackName + "', FinishDate = '" + DateTime.Now + "' where id = '" + KittingId + "'");
            Resp resp = new Resp();
            if (count > 0 && ct > 0)
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

        [System.Web.Http.HttpPost]
        public Resp multipleinrack(itemList list)
        {
            foreach (var item in list.arr)
            {
                SqlHelper.ExecuteNonQuery("update MaterialKittingRack set kittingid = '" + item.kittingId + "' where id = '" + item.rackId + "'");
                SqlHelper.ExecuteNonQuery("update MaterialKitting set RackName = '" + item.RackName + "' where id = '" + item.kittingId + "'");
            }
            Resp resp = new Resp();
            try
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "执行成功";

            }
            catch (Exception)
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "执行失败";
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp outRack(int kittingId)
        {
            int count = SqlHelper.ExecuteNonQuery("update MaterialKittingRack set kittingid = NULL where kittingid = '" + kittingId + "'");
            Resp resp = new Resp();
            if (count > 0)
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
        public Resp getRack()
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select a.*, b.* FROM (select * from MaterialKittingRack)a left join (select * from MaterialKitting where IsDel = 0)b on a.kittingid = b.id");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getRackById(int id)
        {
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select a.*, b.* FROM (select * from MaterialKittingRack)a left join (select * from MaterialKitting where IsDel = 0)b on a.kittingid = b.id where a.id = '" + id + "'");
            Resp resp = new Resp();
            resp.Code = 200;
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp addRack(Rack model)
        {
            int count = SqlHelper.ExecuteNonQuery("insert into MaterialKittingRack (RackName) values ('" + model.RackName + "')");
            Resp resp = new Resp();
            if (count > 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "add success";
            }
            else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "add fail";
            }
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updateRack(Rack model)
        {
            int count = SqlHelper.ExecuteNonQuery("update MaterialKittingRack set RackName = '" + model.RackName + "'");
            Resp resp = new Resp();
            if (count > 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "update success";
            }
            else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "update fail";
            }

            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp deleteRack(int id)
        {
            int count = SqlHelper.ExecuteNonQuery("delete from materialkittingrack where id = '" + id + "'");
            Resp resp = new Resp();
            if (count > 0)
            {
                resp.Code = 200;
                resp.Data = null;
                resp.Message = "delete success";
            }
            else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "delete fail";
            }
            return resp;
        }
    }

    public class itemList
    {
        public List<rackList> arr { get; set; }
    }

    public class rackList
    {
        public int rackId { get; set; }
        public string RackName { get; set; }
        public int kittingId { get; set; }
    }

    public class Resp
    {
        public int Code { get; set; }
        public DataTable Data { get; set; }
        public string Message { get; set; }
    }


    public class Rack
    {
        public string RackName { get; set; }
    }

}
