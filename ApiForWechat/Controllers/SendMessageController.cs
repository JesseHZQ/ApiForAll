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
    public class SendMessageController : ApiController
    {
        const string SendMessageUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=";
        // 任务分配应用的Id
        const int agentid = 1000003;

        [HttpGet]
        public SendMessageResponse SendTextMessage(TextMessage textMessage)
        {
            //TextMessage textMessage = new TextMessage()
            //{
            //    touser = touser,
            //    agentid = agentid,
            //    msgtype = "text",
            //    text = new Txt
            //    {
            //        content = "你的快递已到，请携带工卡前往邮件中心领取。\n出发前可查看<a href=\"http://work.weixin.qq.com\">邮件中心视频实况</a>，聪明避开排队。"
            //    }
            //};
            string AccessToken = new AccessTokenController().GetAccessToken().access_token;
            return JsonConvert.DeserializeObject<SendMessageResponse>(MyAjax.HttpPost(SendMessageUrl + AccessToken, JsonConvert.SerializeObject(textMessage)));
        }

        [HttpGet]
        public SendMessageResponse SendTextCardMessage(TextCardMessage textCardMessage)
        {
            //TextCardMessage textMessage = new TextCardMessage()
            //{
            //    touser = touser,
            //    agentid = agentid,
            //    msgtype = "textcard",
            //    textcard = new TextCard
            //    {
            //        title = "领奖通知",
            //        description = "<div class=\"gray\">" + DateTime.Now + "</div> <div class=\"normal\">恭喜你抽中iPhone 11一台，领奖码：xxxx</div><div class=\"highlight\">请于2019年10月15日前联系行政同事领取</div>",
            //        url = "https://www.baidu.com",
            //        btntxt = "立即领奖"
            //    }
            //};
            string AccessToken = new AccessTokenController().GetAccessToken().access_token;
            return JsonConvert.DeserializeObject<SendMessageResponse>(MyAjax.HttpPost(SendMessageUrl + AccessToken, JsonConvert.SerializeObject(textCardMessage)));
        }
    }
}
