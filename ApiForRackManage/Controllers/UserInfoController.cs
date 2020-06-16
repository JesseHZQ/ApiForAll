using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForRackManage.Controllers
{
    public class UserInfoController : ApiController
    {
        RackDataContext dc = new RackDataContext();

        /// <summary>
        /// 验证用户名和密码
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp checkUser(User user)
        {
            Resp resp = new Resp();
            User item = dc.User.FirstOrDefault(x => x.Name == user.Name);
            if (item != null)
            {
                if (item.PW == user.PW)
                {
                    resp.Code = 200;
                    resp.Data = item;
                    resp.Message = "登录成功";
                }
                else
                {
                    resp.Code = 10001;
                    resp.Data = null;
                    resp.Message = "密码错误";
                }
            }
            else
            {
                resp.Code = 10001;
                resp.Data = null;
                resp.Message = "用户名不存在";
            }
            return resp;
        }
        public class Resp
        {
            public int Code { get; set; }
            public User Data { get; set; }
            public string Message { get; set; }
        }
    }
}
