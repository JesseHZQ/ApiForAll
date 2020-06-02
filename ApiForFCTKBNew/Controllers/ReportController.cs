using ApiForFCTKBNew.Models;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace ApiForFCTKBNew.Controllers
{
    public class ReportController : ApiController
    {
        [HttpGet]
        public string generateExcel(string type)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\UFlex System Status.xls");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 2;
            List<SlotPlan> list = new SlotPlanController().GetSystemByType1(type);
            foreach (SlotPlan s in list)
            {
                for (int i = 0; i < 18; i++)
                {
                    Style sty = cells[y, i].GetStyle();
                    sty.Font.Size = 8;
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cells[y, i].SetStyle(sty);
                }
                cells[y, 0].PutValue(s.PD);
                cells[y, 1].PutValue(s.PlanShipDate);
                cells[y, 2].PutValue(s.Slot);
                cells[y, 3].PutValue(s.Model);
                cells[y, 4].PutValue(s.Customer);
                cells[y, 5].PutValue(s.SO);
                cells[y, 6].PutValue(s.PO);
                cells[y, 14].PutValue(s.Launch);
                cells[y, 15].PutValue(s.CORC);
                List<SlotShortage> listShortage = s.ShortageList;
                string Shortage = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage = Shortage + e.PN + "---" + e.QTY + ";";
                    }
                }
                cells[y, 16].PutValue(Shortage);
                List<SlotEIssue> listEIssue = s.EIssueList;
                string Engineering_Issue = "";
                if (listEIssue.Count > 0)
                {
                    foreach (SlotEIssue e in listEIssue)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item;
                    }
                }
                cells[y, 17].PutValue(Engineering_Issue);

                if (s.Pack != "" && s.Pack != null)
                {
                    cells[y, 7].PutValue(s.Pack);
                    SetCellStyle(cells[y, 7], Shortage, Engineering_Issue);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cells[y, 8].PutValue(s.BU);
                    SetCellStyle(cells[y, 8], Shortage, Engineering_Issue);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cells[y, 9].PutValue(s.QFAA);
                    SetCellStyle(cells[y, 9], Shortage, Engineering_Issue);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cells[y, 10].PutValue(s.CSW);
                    SetCellStyle(cells[y, 10], Shortage, Engineering_Issue);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cells[y, 11].PutValue(s.TestBU);
                    SetCellStyle(cells[y, 11], Shortage, Engineering_Issue);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cells[y, 12].PutValue(s.OI);
                    SetCellStyle(cells[y, 12], Shortage, Engineering_Issue);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cells[y, 13].PutValue(s.CoreBU);
                    SetCellStyle(cells[y, 13], Shortage, Engineering_Issue);
                }
                y++;
            }

            Cells cellsShip = wb.Worksheets[1].Cells;
            int yShip = 2;
            List<SlotPlan> listShip = new SlotPlanController().GetShippedSystemByType(type);
            foreach (SlotPlan s in listShip)
            {
                for (int i = 0; i < 18; i++)
                {
                    Style sty = cellsShip[yShip, i].GetStyle();
                    sty.Font.Size = 8;
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellsShip[yShip, i].SetStyle(sty);
                }
                cellsShip[yShip, 0].PutValue(s.PD);
                cellsShip[yShip, 1].PutValue(s.PlanShipDate);
                cellsShip[yShip, 2].PutValue(s.Slot);
                cellsShip[yShip, 3].PutValue(s.Model);
                cellsShip[yShip, 4].PutValue(s.Customer);
                cellsShip[yShip, 5].PutValue(s.SO);
                cellsShip[yShip, 6].PutValue(s.PO);
                cellsShip[yShip, 14].PutValue(s.Launch);
                cellsShip[yShip, 15].PutValue(s.CORC);
                List<SlotShortage> listShortage = s.ShortageList;
                string Shortage = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage = Shortage + e.PN + "---" + e.QTY + ";";
                    }
                }
                cellsShip[yShip, 16].PutValue(Shortage);
                List<SlotEIssue> listEIssue = s.EIssueList;
                string Engineering_Issue = "";
                if (listEIssue.Count > 0)
                {
                    foreach (SlotEIssue e in listEIssue)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item;
                    }
                }
                cellsShip[yShip, 17].PutValue(Engineering_Issue);

                if (s.Pack != "" && s.Pack != null)
                {
                    cellsShip[yShip, 7].PutValue(s.Pack);
                    SetCellStyle(cellsShip[yShip, 7], Shortage, Engineering_Issue);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cellsShip[yShip, 8].PutValue(s.BU);
                    SetCellStyle(cellsShip[yShip, 8], Shortage, Engineering_Issue);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cellsShip[yShip, 9].PutValue(s.QFAA);
                    SetCellStyle(cellsShip[yShip, 9], Shortage, Engineering_Issue);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cellsShip[yShip, 10].PutValue(s.CSW);
                    SetCellStyle(cellsShip[yShip, 10], Shortage, Engineering_Issue);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cellsShip[yShip, 11].PutValue(s.TestBU);
                    SetCellStyle(cellsShip[yShip, 11], Shortage, Engineering_Issue);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cellsShip[yShip, 12].PutValue(s.OI);
                    SetCellStyle(cellsShip[yShip, 12], Shortage, Engineering_Issue);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cellsShip[yShip, 13].PutValue(s.CoreBU);
                    SetCellStyle(cellsShip[yShip, 13], Shortage, Engineering_Issue);
                }
                yShip++;
            }
            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\SR_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond.ToString() + new Random().Next(1, 10000).ToString() + ".xlsx";
            wb.Save(filePath);
            return filePath;
        }
        
        [HttpGet]
        public string generateExcelWithGroup(string type)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\UFlex System Status With Group.xls");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 2;
            List<SlotPlan> list = new SlotPlanController().GetSystemByType(type);
            foreach (SlotPlan s in list)
            {
                for (int i = 0; i < 19; i++)
                {
                    Style sty = cells[y, i].GetStyle();
                    sty.Font.Size = 8;
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cells[y, i].SetStyle(sty);
                }
                cells[y, 0].PutValue(s.GroupNum);
                cells[y, 1].PutValue(s.PD);
                cells[y, 2].PutValue(s.PlanShipDate);
                cells[y, 3].PutValue(s.Slot);
                cells[y, 4].PutValue(s.Model);
                cells[y, 5].PutValue(s.Customer);
                cells[y, 6].PutValue(s.SO);
                cells[y, 7].PutValue(s.PO);
                cells[y, 15].PutValue(s.Launch);
                cells[y, 16].PutValue(s.CORC);
                List<SlotShortage> listShortage = s.ShortageList;
                string Shortage = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage = Shortage + e.PN + "---" + e.QTY + ";";
                    }
                }
                cells[y, 17].PutValue(Shortage);
                List<SlotEIssue> listEIssue = s.EIssueList;
                string Engineering_Issue = "";
                if (listEIssue.Count > 0)
                {
                    foreach (SlotEIssue e in listEIssue)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item;
                    }
                }
                cells[y, 18].PutValue(Engineering_Issue);

                if (s.Pack != "" && s.Pack != null)
                {
                    cells[y, 8].PutValue(s.Pack);
                    SetCellStyle(cells[y, 8], Shortage, Engineering_Issue);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cells[y, 9].PutValue(s.BU);
                    SetCellStyle(cells[y, 9], Shortage, Engineering_Issue);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cells[y, 10].PutValue(s.QFAA);
                    SetCellStyle(cells[y, 10], Shortage, Engineering_Issue);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cells[y, 11].PutValue(s.CSW);
                    SetCellStyle(cells[y, 11], Shortage, Engineering_Issue);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cells[y, 12].PutValue(s.TestBU);
                    SetCellStyle(cells[y, 12], Shortage, Engineering_Issue);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cells[y, 13].PutValue(s.OI);
                    SetCellStyle(cells[y, 13], Shortage, Engineering_Issue);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cells[y, 14].PutValue(s.CoreBU);
                    SetCellStyle(cells[y, 14], Shortage, Engineering_Issue);
                }
                y++;
            }

            Cells cellsShip = wb.Worksheets[1].Cells;
            int yShip = 2;
            List<SlotPlan> listShip = new SlotPlanController().GetShippedSystemByType(type);
            foreach (SlotPlan s in listShip)
            {
                for (int i = 0; i < 19; i++)
                {
                    Style sty = cellsShip[yShip, i].GetStyle();
                    sty.Font.Size = 8;
                    sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                    cellsShip[yShip, i].SetStyle(sty);
                }
                cellsShip[yShip, 0].PutValue(s.GroupNum);
                cellsShip[yShip, 1].PutValue(s.PD);
                cellsShip[yShip, 2].PutValue(s.PlanShipDate);
                cellsShip[yShip, 3].PutValue(s.Slot);
                cellsShip[yShip, 4].PutValue(s.Model);
                cellsShip[yShip, 5].PutValue(s.Customer);
                cellsShip[yShip, 6].PutValue(s.SO);
                cellsShip[yShip, 7].PutValue(s.PO);
                cellsShip[yShip, 15].PutValue(s.Launch);
                cellsShip[yShip, 16].PutValue(s.CORC);
                List<SlotShortage> listShortage = s.ShortageList;
                string Shortage = "";
                if (listShortage.Count > 0)
                {
                    foreach (SlotShortage e in listShortage)
                    {
                        Shortage = Shortage + e.PN + "---" + e.QTY + ";";
                    }
                }
                cellsShip[yShip, 17].PutValue(Shortage);
                List<SlotEIssue> listEIssue = s.EIssueList;
                string Engineering_Issue = "";
                if (listEIssue.Count > 0)
                {
                    foreach (SlotEIssue e in listEIssue)
                    {
                        Engineering_Issue = Engineering_Issue + e.Item;
                    }
                }
                cellsShip[yShip, 18].PutValue(Engineering_Issue);

                if (s.Pack != "" && s.Pack != null)
                {
                    cellsShip[yShip, 8].PutValue(s.Pack);
                    SetCellStyle(cellsShip[yShip, 8], Shortage, Engineering_Issue);
                }
                else if (s.BU != "" && s.BU != null)
                {
                    cellsShip[yShip, 9].PutValue(s.BU);
                    SetCellStyle(cellsShip[yShip, 9], Shortage, Engineering_Issue);
                }
                else if (s.QFAA != "" && s.QFAA != null)
                {
                    cellsShip[yShip, 10].PutValue(s.QFAA);
                    SetCellStyle(cellsShip[yShip, 10], Shortage, Engineering_Issue);
                }
                else if (s.CSW != "" && s.CSW != null)
                {
                    cellsShip[yShip, 11].PutValue(s.CSW);
                    SetCellStyle(cellsShip[yShip, 11], Shortage, Engineering_Issue);
                }
                else if (s.TestBU != "" && s.TestBU != null)
                {
                    cellsShip[yShip, 12].PutValue(s.TestBU);
                    SetCellStyle(cellsShip[yShip, 12], Shortage, Engineering_Issue);
                }
                else if (s.OI != "" && s.OI != null)
                {
                    cellsShip[yShip, 13].PutValue(s.OI);
                    SetCellStyle(cellsShip[yShip, 13], Shortage, Engineering_Issue);
                }
                else if (s.CoreBU != "" && s.CoreBU != null)
                {
                    cellsShip[yShip, 14].PutValue(s.CoreBU);
                    SetCellStyle(cellsShip[yShip, 14], Shortage, Engineering_Issue);
                }
                yShip++;
            }
            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\SR_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond + ".xlsx";
            wb.Save(filePath);
            return filePath;
        }
        
        [HttpGet]
        public HttpResponseMessage downloadFile(string filePath, string fileName)
        {
            try
            {
                var stream = new FileStream(filePath, FileMode.Open);
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = fileName
                };
                return response;
            }
            catch
            {
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
        }

        public void SetCellStyle(Cell cell, string shortage, string issue)
        {
            Style sty = cell.GetStyle();
            if (shortage != "" || issue != "")
            {
                sty.BackgroundColor = System.Drawing.Color.Red;
                sty.Pattern = BackgroundType.Solid;
                sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
            }
            else
            {
                sty.BackgroundColor = System.Drawing.Color.Green;
                sty.Pattern = BackgroundType.Solid;
                sty.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                sty.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
            }
            cell.SetStyle(sty);
        }
    }
}
