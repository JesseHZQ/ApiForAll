using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ApiForWechat.Models;
using Dapper;

namespace ApiForWechat.Controllers
{
    public class SendAssignWorkMessageController : ApiController
    {
        const int agentid = 1000003;
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ControlTower"].ConnectionString);

        [HttpGet]
        public List<SendMessageResponse> SendAssignWorkMessage()
        {
            List<SendMessageResponse> listResponse = new List<SendMessageResponse>();
            string sql = "select * from CT_LEDMaster a left join CT_User b on a.OwnerId = b.UserId";
            List<LED> list = conn.Query<LED>(sql).ToList().Where(x => x.WeChatName != "" && x.WeChatName != null).ToList();
            foreach (LED led in list)
            {
                TextCardMessage textCardMessage = new TextCardMessage
                {
                    touser = led.WeChatName,
                    agentid = agentid,
                    msgtype = "textcard",
                    textcard = new TextCard
                    {
                        title = "任务分配通知",
                        description = "<div class=\"gray\">" + DateTime.Now + "</div> <div class=\"normal\">Hi, " + led.UserName + ". You are assigned system " + led.SystemName + " at station " + led.Station + " in Bay" + led.BayName + ".</div><div class=\"highlight\">FCT Testing</div>",
                        url = "https://www.baidu.com",
                        btntxt = "加油"
                    }
                };
                listResponse.Add(new SendMessageController().SendTextCardMessage(textCardMessage));
            }
            return listResponse;
        }
    }
}
