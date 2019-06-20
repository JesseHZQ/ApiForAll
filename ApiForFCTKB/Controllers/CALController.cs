using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ApiForFCTKB.Models;
using Aspose.Cells;

namespace ApiForFCTKB.Controllers
{
    public class CALController : ApiController
    {
        [HttpGet]
        public string GenerateCSV()
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            List<CAL> list = GetCSVData();
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;
            cells[0, 0].PutValue("工具型号");
            cells[0, 1].PutValue("工具编号");
            cells[0, 2].PutValue("定义值");
            cells[0, 3].PutValue("公差±5% ");
            cells[0, 4].PutValue("公差±10% ");
            cells[0, 5].PutValue("校准设备编号");
            cells[0, 6].PutValue("校准前");
            cells[0, 7].PutValue("校准后1");
            cells[0, 8].PutValue("校准后2");
            cells[0, 9].PutValue("校准后3");
            cells[0, 10].PutValue("第一次校准后平均值(±5%)");
            cells[0, 11].PutValue("第二次校准后1");
            cells[0, 12].PutValue("第二次校准后2");
            cells[0, 13].PutValue("第二次校准后3");
            cells[0, 14].PutValue("第二次校准后平均值(±5%)");
            cells[0, 15].PutValue("漏电电压<20mV");
            cells[0, 16].PutValue("接地电阻值<20Ω");
            cells[0, 17].PutValue("接地电阻值<1012Ω");

            int row = 0;
            foreach (CAL cal in list)
            {
                row++;
                cells[row, 0].PutValue(cal.ToolName);
                cells[row, 1].PutValue(cal.ToolCode);
                cells[row, 2].PutValue(cal.DefinedVal);
                cells[row, 3].PutValue(cal.GX_5);
                cells[row, 4].PutValue(cal.GX_10);
                cells[row, 5].PutValue(cal.EquipCode);
                cells[row, 6].PutValue(cal.X1);
                cells[row, 7].PutValue(cal.X2);
                cells[row, 8].PutValue(cal.X3);
                cells[row, 9].PutValue(cal.X4);
                cells[row, 10].PutValue(cal.Average);
                cells[row, 11].PutValue(cal.N1);
                cells[row, 12].PutValue(cal.N2);
                cells[row, 13].PutValue(cal.N3);
                cells[row, 14].PutValue(cal.N4);
                cells[row, 15].PutValue(cal.N5);
                cells[row, 16].PutValue(cal.N6);
                cells[row, 17].PutValue(cal.N7);
            }

            wb.Save(@"\\suznt004\TER\Mechanical PE\CAL Software database(No deletion)\result.xlsx");
            return "OK";
        }


        public List<CAL> GetCSVData()
        {
            List<CAL> list = new List<CAL>();
            string path = @"\\10.194.51.14\TER\Mechanical PE\CAL Software database(No deletion)\RowData";
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                path = file.FullName;
                Workbook wb = new Workbook(path);
                Cells cells = wb.Worksheets[0].Cells;
                CAL cal = new CAL();
                cal.ToolName = cells[6, 0].Value.ToString().Substring(cells[6, 0].Value.ToString().LastIndexOf(' ') + 1);
                cal.ToolCode = cal.ToolName;
                cal.DefinedVal = cells[10, 0].Value.ToString().Substring(15).Trim();
                float GX = float.Parse(cal.DefinedVal.Split(' ')[0]);
                cal.GX_5 = Math.Round(GX * 0.95, 2) + "-" + Math.Round(GX * 1.05, 2);
                cal.GX_10 = Math.Round(GX * 0.9, 2) + "-" + Math.Round(GX * 1.1, 2);
                cal.EquipCode = cells[7, 0].Value.ToString().Substring(cells[6, 0].Value.ToString().LastIndexOf(' ') + 1);
                cal.X1 = cells[32, 1].Value == null ? "0" : cells[32, 1].Value.ToString();
                cal.X2 = cells[33, 1].Value == null ? "0" : cells[33, 1].Value.ToString();
                cal.X3 = cells[34, 1].Value == null ? "0" : cells[34, 1].Value.ToString();
                cal.X4 = cells[35, 1].Value == null ? "0" : cells[35, 1].Value.ToString();
                cal.Average = decimal.Round((decimal.Parse(cal.X2) + decimal.Parse(cal.X3) + decimal.Parse(cal.X4)) / 3, 2).ToString();
                cal.N1 = "N/A";
                cal.N2 = "N/A";
                cal.N3 = "N/A";
                cal.N4 = "N/A";
                cal.N5 = "N/A";
                cal.N6 = "N/A";
                cal.N7 = "N/A";
                list.Add(cal);
            }
            return list;
        }
    }
}
