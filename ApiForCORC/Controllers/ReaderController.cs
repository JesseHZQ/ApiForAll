using ApiForCORC.Models;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForCORC.Controllers
{
    public class ReaderController : ApiController
    {
        [HttpGet]
        public List<CORC> GetInfo()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");

            List<CORC> List = new List<CORC>();
            List<CORCFilePath> CORCFilePathList = GetFilePathList();
            foreach (CORCFilePath item in CORCFilePathList)
            {
                CORC corc = new CORC();
                corc.SO = item.SO;
                corc.CorcInfo = GetCORCInfo(item);
                corc.SoInfo = GetSOInfo(item);
                List.Add(corc);
            }
            return List;
        }
        public CORCInfo GetCORCInfo(CORCFilePath corcPath)
        {
            CORCInfo corcInfo = new CORCInfo();
            corcInfo.SO = corcPath.SO;
            string[] Lines = File.ReadAllLines(corcPath.CORCPath);
            string LinesConcat = string.Join("", Lines);
            corcInfo.Status = MidStrEx(LinesConcat, "Status: ", "General");
            corcInfo.Customer = MidStrEx(LinesConcat, "Customer		", "	Division");
            corcInfo.SO_Verify = MidStrEx(LinesConcat, "Sales Order #	", "Sales Admin");
            corcInfo.ProductComments = MidStrEx(LinesConcat, "Product comments:", "Reg. Sales Mgr");
            corcInfo.FactoryRequirement = MidStrEx(LinesConcat, "Factory Requirements		", "Site Requirements");
            corcInfo.SiteVoltage = int.Parse(MidStrEx(LinesConcat, "Site Voltage Requirements	", "UltraFLEX System DUT Orientation"));
            corcInfo.DUTOrientation = MidStrEx(LinesConcat, "Orientation must be selected before system shipment.", "6/19/2017:");
            corcInfo.ComputerEquipped = MidStrEx(LinesConcat, "There is no OS installed on the HDD drive.	", "Hard Drive for IGXL set up:");
            corcInfo.IGXLVersion = MidStrEx(LinesConcat, "please detail the need in the text field below.", "Windows XP is no longer available ");
            corcInfo.IPAddress = MidStrEx(LinesConcat, "IP Names and Addresses	", "	Other Specials");
            corcInfo.OtherSpecial = MidStrEx(LinesConcat, "Other Specials	", "	Special Comments And Miscellaneous Attachments:");
            //corcInfo.SpecialComments = MidStrEx(LinesConcat, "Factory Requirements		", "Site Requirements");
            return corcInfo;
        }
        public SOInfo GetSOInfo(CORCFilePath corcPath)
        {
            SOInfo soInfo = new SOInfo();
            soInfo.SO = corcPath.SO;
            LoadOptions loadOptions = new LoadOptions(LoadFormat.TSV);
            Workbook wb = new Workbook(corcPath.SOPath, loadOptions);
            Cells cells = wb.Worksheets[0].Cells;
            DataTable dt = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
            soInfo.SOLineList = new List<SOLine>();
            foreach (DataRow item in dt.Rows)
            {
                SOLine soLine = new SOLine();
                soLine.Line = item["Line"].ToString();
                soLine.OrderedItem = item["Ordered Item"].ToString();
                soLine.Qty = item["Qty"].ToString();
                soLine.QtyShipped = item["Qty Shipped"].ToString();
                soLine.ItemDescription = item["Item Description"].ToString();
                soLine.ScheduleShipDate = item["Schedule Ship Date"].ToString();
                soInfo.SOLineList.Add(soLine);
            }
            return soInfo;
        }
        public List<CORCFilePath> GetFilePathList()
        {
            List<CORCFilePath> List = new List<CORCFilePath>();
            string RootPath = @"C:\Users\suzjeshe\Desktop\CORC";
            string[] DirSO = Directory.GetDirectories(RootPath);
            foreach (string path in DirSO)
            {
                CORCFilePath corcPath = new CORCFilePath();
                corcPath.SO = path.Substring(path.LastIndexOf('\\') + 1);
                string[] Dir = Directory.GetFiles(path);
                foreach (string file in Dir)
                {
                    if (file.Contains("SO"))
                    {
                        corcPath.SOPath = file;
                    }
                    if (file.Contains("ASCII"))
                    {
                        corcPath.CORCPath = file;
                    }
                    if (file.Contains("LOC"))
                    {
                        corcPath.LocatorPath = file;
                    }
                }
                List.Add(corcPath);
            }
            return List;
        }

        /// <summary>
        /// 截取两字符串之间
        /// </summary>
        /// <param name="sourse"></param>
        /// <param name="startstr"></param>
        /// <param name="endstr"></param>
        /// <returns></returns>
        public string MidStrEx(string sourse, string startstr, string endstr)
        {
            string result = string.Empty;
            int startindex, endindex;
            try
            {
                startindex = sourse.IndexOf(startstr);
                if (startindex == -1)
                    return result;
                string tmpstr = sourse.Substring(startindex + startstr.Length);
                endindex = tmpstr.IndexOf(endstr);
                if (endindex == -1)
                    return result;
                result = tmpstr.Remove(endindex);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return result;
        }
    }
}
