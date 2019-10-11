using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForWechat.Models
{
    public class CommonMessageFormat
    {
        // touser、toparty、totag不能同时为空
        // 非必填 成员ID列表（消息接收者，多个接收者用‘|’分隔，最多支持1000个）。特殊情况：指定为 @all，则向该企业应用的全部成员发送
        public string touser { get; set; }
        // 非必填 部门ID列表，多个接收者用‘|’分隔，最多支持100个。当touser为@all时忽略本参数
        public string toparty { get; set; }
        // 非必填 标签ID列表，多个接收者用‘|’分隔，最多支持100个。当touser为@all时忽略本参数
        public string totag { get; set; }
        // 必填 消息类型
        public string msgtype { get; set; }
        // 必填 企业应用的id，整型。企业内部开发，可在应用的设置页面查看；第三方服务商，可通过接口 获取企业授权信息 获取该参数值
        public int agentid { get; set; }
        // 非必填 表示是否是保密消息，0表示否，1表示是，默认0
        public int safe { get; set; }
        // 非必填 表示是否开启id转译，0表示否，1表示是，默认0
        public int enable_id_trans { get; set; }
    }
}