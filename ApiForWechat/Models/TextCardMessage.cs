using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForWechat.Models
{
    public class TextCardMessage: CommonMessageFormat
    {
        public TextCard textcard { get; set; }
    }

    public class TextCard
    {
        public string title { get; set; }
        public string description { get; set; }
        public string url { get; set; }
        public string btntxt { get; set; }
    }
}