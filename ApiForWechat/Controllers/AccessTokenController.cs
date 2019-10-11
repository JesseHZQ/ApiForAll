using ApiForWechat.HttpRequest;
using ApiForWechat.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForWechat.Controllers
{
    public class AccessTokenController : ApiController
    {
        const string corpid = "ww9c471f1d0a3fd2d6";
        // 任务分配应用的 secret
        const string corpsecret = "aRUAoVIAEykUMaYxYVph8zZwJ1AtELUZjsLdNeV9EE0";
        const string AccessTokenUrl = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret;

        [HttpGet]
        public AccessToken GetAccessToken()
        {
            string str = MyAjax.HttpGet(AccessTokenUrl);
            return JsonConvert.DeserializeObject<AccessToken>(str);
        }
    }
}
