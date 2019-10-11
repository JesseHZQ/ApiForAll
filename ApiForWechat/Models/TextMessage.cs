using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForWechat.Models
{
    public class TextMessage: CommonMessageFormat
    {
        public Txt text { get; set; }
    }

    public class Txt
    {
        public string content { get; set; }
    }
}