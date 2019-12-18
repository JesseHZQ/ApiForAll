using ApiForCORC.Models;
using Aspose.Cells;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web.Http;

namespace ApiForCORC.Controllers
{
    public class ExcelSOController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SlotKB"].ConnectionString);
        [HttpGet]
        public string GetExcelSO()
        {
            string sqlQuery = "select * from KANBAN_SLOTPLAN";
            List<SlotPlan> SlotPlanList = conn.Query<SlotPlan>(sqlQuery).ToList();
            string path = @"\\10.194.51.14\TER\FCT E-KANBAN Database\SalesOrder";
            string flexPath = "";
            string JPath = "";
            string MiscPath = "";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (file.FullName.Contains("~"))
                {
                    continue;
                }
                if (file.FullName.Contains("Flex Family"))
                {
                    flexPath = file.FullName;
                }
                if (file.FullName.Contains("J750 Family"))
                {
                    JPath = file.FullName;
                }
                if (file.FullName.Contains("Misc Order"))
                {
                    MiscPath = file.FullName;
                }
            }
            List<ExcelSO> list = new List<ExcelSO>();


            // Flex
            Workbook wb = new Workbook(flexPath);
            Cells cells = wb.Worksheets[0].Cells;
            DataTable dt = cells.ExportDataTableAsString(2, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                string status = item["Issues"].ToString();
                if (status == "YES")
                {
                    ExcelSO info = new ExcelSO();
                    info.Slot = item["System #"].ToString();
                    info.SO = item["Teradyne S/O #"].ToString();
                    info.Customer = item["Customer Name"].ToString();
                    info.ShipToLocation = item["Ship To Location"].ToString();
                    info.Incoterms = item["Incoterms"].ToString();
                    info.Week = item["Week#"].ToString().Substring(4, 2);
                    info.Issues = item["Comments"].ToString();
                    list.Add(info);
                }
            }

            // J750
            wb = new Workbook(JPath);
            cells = wb.Worksheets[0].Cells;
            dt = cells.ExportDataTableAsString(2, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                string status = item["Issues"].ToString();
                if (status == "YES")
                {
                    ExcelSO info = new ExcelSO();
                    info.Slot = item["System #"].ToString();
                    info.SO = item["Teradyne S/O #"].ToString();
                    info.Customer = item["Customer Name"].ToString();
                    info.ShipToLocation = item["Ship To Location"].ToString();
                    info.Incoterms = item["Incoterms"].ToString();
                    info.Week = item["Week #"].ToString().Substring(4, 2);
                    info.Issues = item["Comments"].ToString();
                    list.Add(info);
                }
            }

            // Misc
            wb = new Workbook(MiscPath);
            cells = wb.Worksheets[0].Cells;
            dt = cells.ExportDataTableAsString(1, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            foreach (DataRow item in dt.Rows)
            {
                string status = item["SO Approved"].ToString();
                if (status != "Y" && status != "")
                {
                    ExcelSO info = new ExcelSO();
                    info.Slot = "NA";
                    info.SO = item["Teradyne SO#"].ToString();
                    info.Customer = item["Customer"].ToString();
                    info.ShipToLocation = item["Ship To Location"].ToString();
                    info.Incoterms = item["Incoterm"].ToString();
                    info.Week = item["Promise Week"].ToString().Substring(4, 2);
                    info.Issues = item["SO Approved"].ToString().Substring(2);
                    list.Add(info);
                }
            }




            // 获取当前日期
            string nowDate = Tool.GetFormatDate(DateTime.Now); // 1948
            string prev1Date = Tool.GetFormatDate(DateTime.Now.AddDays(-7)); // 1947
            string prev2Date = Tool.GetFormatDate(DateTime.Now.AddDays(-14)); // 1946
            string next1Date = Tool.GetFormatDate(DateTime.Now.AddDays(7)); // 1949
            string next2Date = Tool.GetFormatDate(DateTime.Now.AddDays(14)); // 1950

            int now = Tool.GetDay(DateTime.Now); // 3
            int next1 = Tool.GetDay(DateTime.Now.AddDays(7)); // 4
            int next2 = Tool.GetDay(DateTime.Now.AddDays(14)); // 5

            foreach (ExcelSO item in list)
            {
                if (nowDate.Substring(2, 2) == item.Week)
                {
                    item.IsShow = true;
                    item.PDWeek = nowDate + ".5";
                    item.DueDateToBeFixed = "WK" + prev1Date + ".5";
                    item.RemainFixedDay = "Over Due";
                    item.OverDueWorkingDay = now.ToString();
                    if (now == 5)
                    {
                        item.EarliestShipDay = "WK" + Tool.GetFormatDateWithDay(DateTime.Now.AddDays(10)).ToString();
                    }
                    item.EarliestShipDay = "WK" + Tool.GetFormatDateWithDay(DateTime.Now.AddDays(8)).ToString();
                    item.Impacts = "Slip to WK" + next1Date;
                }
                if (next1Date.Substring(2, 2) == item.Week)
                {
                    item.IsShow = true;
                    item.PDWeek = next1Date + ".5";
                    item.DueDateToBeFixed = "WK" + nowDate + ".5";
                    if (now == 5)
                    {
                        item.RemainFixedDay = "Due";
                        item.OverDueWorkingDay = "0";
                        item.EarliestShipDay = "WK" + Tool.GetFormatDateWithDay(DateTime.Now.AddDays(10)).ToString();
                        item.Impacts = "Slip to WK" + next2Date;
                    }
                    else
                    {
                        item.RemainFixedDay = (5 - now).ToString();
                        item.OverDueWorkingDay = "NA";
                        item.EarliestShipDay = "WK" + Tool.GetFormatDateWithDay(DateTime.Now.AddDays(8)).ToString();
                        item.Impacts = "";
                    }
                }
            }
            list = list.Where(x => x.IsShow && !x.Slot.Contains("C")).ToList();

            //string sql = "insert into KANBAN_CORC (Slot, SO, MRPWeek, CORCStatus, CORCDueDateToBeFixed, RemainFixedDay, OverDueWorkingDay, EarliestShipDay, Impacts, InsertTime) values (@Slot, @SO, @MRPWeek, @CORCStatus, @CORCDueDateToBeFixed, @RemainFixedDay, @OverDueWorkingDay, @EarliestShipDay, @Impacts, @InsertTime)";
            //conn.Execute(sql, list);

            MailSender ms = new MailSender();
            string date = DateTime.Now.ToString("dddd, MMM d", DateTimeFormatInfo.InvariantInfo);
            ms.Subject = string.Format("SO Alert on {0}", date);
            ms.From = "SO Alert SO-Alert@flex.com";
            ms.ToList = new List<string>();
            ms.ToList.Add("amanda_niu@notes.teradyne.com");
            ms.ToList.Add("veda_ma@notes.teradyne.com");
            ms.ToList.Add("ruby_wang@notes.teradyne.com");
            ms.ToList.Add("napo_zhang@notes.teradyne.com");
            ms.ToList.Add("napo.zhang@flex.com");
            ms.ToList.Add("robert.yu@flex.com");
            ms.ToList.Add("aivin.shao@flex.com");
            ms.ToList.Add("jane.sha@flex.com");
            ms.ToList.Add("luna.wang@flex.com");
            ms.ToList.Add("jesse.he@flex.com");
            ms.CCList = new List<string>();
            ms.FilePathList = new List<string>();
            ms.FilePathList.Add(GenerateExcel(list));
            ms.Content = GetMailContent(list);
            return SendMail(ms);
        }

        public string GenerateExcel(List<ExcelSO> l)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\SO Alert Template.xlsx");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 3;
            List<ExcelSO> list = l.OrderBy(x => x.PDWeek).ToList();
            foreach (ExcelSO s in list)
            {
                for (int i = 0; i < 12; i++)
                {
                    var sty = cells[y, i].GetStyle();
                    sty.Pattern = BackgroundType.Solid;
                    if (s.RemainFixedDay == "Over Due")
                    {
                        sty.ForegroundColor = Color.Red;
                    }
                    else if (s.RemainFixedDay == "Due")
                    {
                        sty.ForegroundColor = Color.Yellow;
                    }
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cells[y, i].SetStyle(sty);
                }

                cells[y, 0].PutValue(s.PDWeek);
                cells[y, 1].PutValue(s.Slot);
                cells[y, 2].PutValue(s.SO);
                cells[y, 3].PutValue(s.Customer);
                cells[y, 4].PutValue(s.ShipToLocation);
                cells[y, 5].PutValue(s.Incoterms);
                cells[y, 6].PutValue(s.Issues);
                cells[y, 7].PutValue(s.DueDateToBeFixed);
                cells[y, 8].PutValue(s.RemainFixedDay);
                cells[y, 9].PutValue(s.OverDueWorkingDay);
                cells[y, 10].PutValue(s.EarliestShipDay);
                cells[y, 11].PutValue(s.Impacts);
                y++;
            }
            cells[y + 2, 0].PutValue("Note:");
            cells[y + 3, 0].PutValue("1.Lead time from hold to released --- 5 working days");
            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\SO_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond.ToString() + new Random().Next(1, 10000).ToString() + ".xlsx";
            wb.Save(filePath);
            return filePath;
        }

        public string GetMailContent(List<ExcelSO> l)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format(@"<p style='color: #000; margin-bottom: 20px;'>Hello Everyone,</p>
              <p style='color: #000; margin-bottom: 20px;'>Here is SO status alert on {0}.</p>
              <table style='table-layout:fixed; border-style: solid; border-width: 1px;text-align: center; border-collapse: collapse; font-size: 12px; line-height: 1;color: #000;' cellspacing='0'>
                <tr style='background-color: #5294f7; font-weight: bold; color: #000; font-size: 14px; height: 24px; text-align: center;'  >
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>PD</div><div style='margin: 0;'>Week</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'>Slot</th>
                  <th width='100' style='border-style: solid; border-width: 1px'>SO</th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Customer</div><div style='margin: 0;'>Name</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Ship to</div><div style='margin: 0;'>Location</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'>Incoterms</th>
                  <th width='100' style='border-style: solid; border-width: 1px'>Issues</th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Due date to</div><div style='margin: 0;'>be fixed</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Remained</div><div style='margin: 0;'>fix day</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Overdue aging</div><div style='margin: 0;'>working day</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'><div style='margin: 0;'>Earliest</div><div style='margin: 0;'>ship day</div></th>
                  <th width='100' style='border-style: solid; border-width: 1px'>Impact</th>
                </tr>", DateTime.Now.ToString("MMM d", DateTimeFormatInfo.InvariantInfo)));
            foreach (ExcelSO item in l.OrderBy(x => x.PDWeek).ToList())
            {
                string bgc = "";
                if (item.RemainFixedDay == "Over Due")
                {
                    bgc = "#f00";
                }
                else if (item.RemainFixedDay == "Due")
                {
                    bgc = "Yellow";
                }
                else
                {
                    bgc = "";
                }
                sb.Append(string.Format(@"<tr style='line-height: 20px; background: {0};'>
                    <td style='border-style: solid; border-width: 1px'>{1}</td>
                    <td style='border-style: solid; border-width: 1px'>{2}</td>
                    <td style='border-style: solid; border-width: 1px'>{3}</td>
                    <td style='border-style: solid; border-width: 1px'>{4}</td>
                    <td style='border-style: solid; border-width: 1px'>{5}</td>
                    <td style='border-style: solid; border-width: 1px;'>{6}</td>
                    <td style='border-style: solid; border-width: 1px; text-align: left'>{7}</td>
                    <td style='border-style: solid; border-width: 1px;'>{8}</td>
                    <td style='border-style: solid; border-width: 1px;'>{9}</td>
                    <td style='border-style: solid; border-width: 1px;'>{10}</td>
                    <td style='border-style: solid; border-width: 1px;'>{11}</td>
                    <td style='border-style: solid; border-width: 1px;'>{12}</td>
                    </tr>", bgc, item.PDWeek, item.Slot, item.SO, item.Customer, item.ShipToLocation, item.Incoterms, item.Issues, item.DueDateToBeFixed, item.RemainFixedDay, item.OverDueWorkingDay, item.EarliestShipDay, item.Impacts));
            }
            sb.Append("</table>");
            sb.Append("<p style='color: #000; font-weight: 700;'>Note: </p>");
            sb.Append("<p style='color: #000;'>1.Lead time from <span style='text-decoration: underline'>hold</span> to <span style='text-decoration: underline'>released</span> --- 5 working days</p>");
            sb.Append("<div style='margin: 20px 0 0;'>Best Regards!</div>");
            sb.Append("<div style='margin: 0 0 50px;'>SO Alert System</div>");
            sb.Append("<p style='font-style: italic;'>Auto mail, please do not reply, thanks!</p>");
            return sb.ToString();
        }

        public string SendMail(MailSender mailSender)
        {
            MailMessage mess = new MailMessage();
            mess.From = new MailAddress(mailSender.From);
            mess.Subject = string.Format(mailSender.Subject);
            mess.IsBodyHtml = true;
            mess.BodyEncoding = System.Text.Encoding.UTF8;
            SmtpClient client = new SmtpClient();
            client.Host = "10.194.51.14";
            client.Credentials = new System.Net.NetworkCredential("root", "Flex_1969");
            foreach (string item in mailSender.ToList)
            {
                mess.To.Add(item);
            }
            if (mailSender.CCList != null)
            {
                foreach (string item in mailSender.CCList)
                {
                    mess.CC.Add(item);
                }
            }
            if (mailSender.FilePathList != null)
            {
                foreach (string item in mailSender.FilePathList)
                {
                    Attachment at = new Attachment(@item, MediaTypeNames.Application.Octet);
                    //at.Name = item.Substring(item.LastIndexOf('\\'));
                    at.Name = "SO_Alert.xlsx";
                    mess.Attachments.Add(at);
                }
            }
            mess.Body = string.Format(mailSender.Content);
            try
            {
                client.Send(mess);
                return "OK";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        public class MailSender
        {
            public string Subject { get; set; }
            public string From { get; set; }
            public List<string> ToList { get; set; }
            public List<string> CCList { get; set; }
            public List<string> FilePathList { get; set; }
            public string Content { get; set; }
        }
    }
}
