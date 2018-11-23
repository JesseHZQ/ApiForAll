using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Globalization;
using CheckSlot.DataContext;

/// <summary>
/// Summary description for MailTask
/// </summary>
public class MailTask
{
    public static void SendMail()
    {
        MailMessage mess = new MailMessage();
        mess.From = new MailAddress("SlotCheck@flex.com");
        mess.Subject = string.Format("校验 wk{0}", GetWeekOfYear(DateTime.Now));
        mess.IsBodyHtml = true;
        mess.BodyEncoding = System.Text.Encoding.UTF8;
        SmtpClient client = new SmtpClient();
        client.Host = "10.194.51.14";
        client.Credentials = new System.Net.NetworkCredential("SMTPUser", "User_001");
        mess.To.Add("Jesse.He@flex.com");
        mess.To.Add("Napo.Zhang@flex.com");
        mess.To.Add("sky.zong@flex.com");
        mess.To.Add("sean.hu@flex.com");
        mess.To.Add("gary.huang@flex.com");
        mess.To.Add("lisa.qian@flex.com");
        mess.To.Add("michelle.li@flex.com");
        mess.To.Add("fengjuan.shen@flex.com");
        mess.To.Add("alex.tao@flex.com");
        mess.To.Add("marcus.meng@flex.com");
        mess.To.Add("jacqueline.ge@flex.com");


        mess.CC.Add("eric.lin2@flex.com");
        mess.CC.Add("liang.fan@flex.com");
        mess.CC.Add("martin.hao@flex.com");
        mess.CC.Add("justin.mu@flex.com");
        mess.CC.Add("eddie.jiang@flex.com");
        mess.CC.Add("evelyn.du@flex.com");
        mess.CC.Add("pan.cheng@flex.com");
        mess.CC.Add("caifeng.zhu@flex.com");
        mess.CC.Add("chris.liu1@flex.com");

        string str = GetMessage();
        mess.Body = string.Format(str);
        client.Send(mess);
    }

    /// <summary>
    /// 获取指定日期，在为一年中为第几周
    /// </summary>
    /// <param name="dt">指定时间</param>
    /// <reutrn>返回第几周</reutrn>
    private static int GetWeekOfYear(DateTime dt)
    {
        GregorianCalendar gc = new GregorianCalendar();
        int weekOfYear = gc.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        return weekOfYear;
    }

    public static string GetMessage()
    {
        SlotDemandDataContext dc = new SlotDemandDataContext();
        List<SlotDemandRequestDto> list = new List<SlotDemandRequestDto>();
        list = dc.SlotDemandRequestDto.Where(
          x => x.ItemNo == "TDN-425-240-00"
            || x.ItemNo == "TDN-640-128-00"
            || x.ItemNo == "TDN-361-749-03"
            || x.ItemNo == "TDN-429-291-01"
            || x.ItemNo == "TDN-361-749-04"
            || x.ItemNo == "TDN-361-749-05"
            || x.ItemNo == "TDN-361-749-08"
        ).ToList();

        string str = @"<table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #4E575B;' cellspacing='0'><thead>";
        str += "<tr style='font-weight: bold; color: #4183C4; font-size: 14px; height: 24px; text-align: center;'>";
        str += "<th style='border-style: solid; border-width: 1px'>PONum</th>";
        str += "<th style='border-style: solid; border-width: 1px'>SLOT</th>";
        str += "<th style='border-style: solid; border-width: 1px'>WK</th>";
        str += "<th style='border-style: solid; border-width: 1px'>Promise Date</th>";
        str += "<th style='border-style: solid; border-width: 1px'>CustomerName</th>";
        str += "<th style='border-style: solid; border-width: 1px'>Component</th>";
        str += "<th style='border-style: solid; border-width: 1px'>Extended Qty</th>";
        str += "<th style='border-style: solid; border-width: 1px'>Comp Description</th></tr>";
        str += "</thead><tbody>";
        for (int i = 0; i < list.Count; i++)
        {
            str += "<tr style='font-size: 14px; height: 24px; text-align: left;'>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].RefDocNo + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].Slot + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].WK + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].RequiredDate + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].Customer + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].ItemNo + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].RequiredQty + "</td>";
            str += "<td style='border-style: solid; border-width: 1px'>" + list[i].ItemDescription + "</td>";
            str += "</tr>";
        }
        str += "</tbody></table>";
        return str;
    }
}