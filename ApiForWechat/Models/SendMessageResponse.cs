using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForWechat.Models
{
    public class SendMessageResponse
    {
        public int errcode { get; set; }
        public string errmsg { get; set; }
        // 不区分大小写，返回的列表都统一转为小写
        public string invaliduser { get; set; }
        public string invalidparty { get; set; }
        public string invalidtag { get; set; }
    }
}