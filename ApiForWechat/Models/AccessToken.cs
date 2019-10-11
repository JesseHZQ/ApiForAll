using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForWechat.Models
{
    public class AccessToken
    {
        // 出错返回码，为0表示成功，非0表示调用失败
        public int errcode { get; set; }
        // 返回码提示语
        public string errmsg { get; set; }
        // 获取到的凭证，最长为512字节
        public string access_token { get; set; }
        // 凭证的有效时间（秒）
        public int expires_in { get; set; }
    }
}