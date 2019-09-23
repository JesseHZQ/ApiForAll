using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace ApiForFCTKB.Controllers
{
    public class AMLController : ApiController
    {
        [HttpPost]
        public string generateExcel(AMLList list)
        {
            // 注册Aspose License
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");
            Workbook wb = new Workbook(@"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\AML Diff PE.xlsx");
            Cells cells = wb.Worksheets[0].Cells;
            int y = 0;
            foreach (AmlDiffPe aml in list.list)
            {
                y++;
                cells[y, 0].PutValue(aml.TeradynePartNumber);
                cells[y, 1].PutValue(aml.Manufacturer);
                cells[y, 2].PutValue(aml.CorpStandardDescAfter);
                cells[y, 3].PutValue(aml.ParentPackageNameAfter);
                cells[y, 4].PutValue(aml.AqueousCompatibleBefore);
                cells[y, 5].PutValue(aml.AqueousCompatibleAfter);
                cells[y, 6].PutValue(aml.AqueousCompatibleChanged);
                cells[y, 7].PutValue(aml.SolventSensitiveBefore);
                cells[y, 8].PutValue(aml.SolventSensitiveAfter);
                cells[y, 9].PutValue(aml.SolventSensitiveChanged);
                cells[y, 10].PutValue(aml.MoistureSensitivityRatingBefore);
                cells[y, 11].PutValue(aml.MoistureSensitivityRatingAfter);
                cells[y, 12].PutValue(aml.MoistureSensitivityRatingChanged);
                cells[y, 13].PutValue(aml.IsRohsCompliantBefore);
                cells[y, 14].PutValue(aml.IsRohsCompliantAfter);
                cells[y, 15].PutValue(aml.IsRohsCompliantChanged);
                cells[y, 16].PutValue(aml.MaxProcTempInDegBefore);
                cells[y, 17].PutValue(aml.MaxProcTempInDegAfter);
                cells[y, 18].PutValue(aml.MaxProcTempInDegChanged);
                cells[y, 19].PutValue(aml.DurationMaxTempSecBefore);
                cells[y, 20].PutValue(aml.DurationMaxTempSecAfter);
                cells[y, 21].PutValue(aml.DurationMaxTempSecChanged);
                cells[y, 22].PutValue(aml.LeadFinishBefore);
                cells[y, 23].PutValue(aml.LeadFinishAfter);
                cells[y, 24].PutValue(aml.LeadFinishChanged);
                cells[y, 25].PutValue(aml.FoundryProcessCommentsBefore);
                cells[y, 26].PutValue(aml.FoundryProcessCommentsAfter);
                cells[y, 27].PutValue(aml.FoundryProcessCommentsChanged);
                cells[y, 28].PutValue(aml.T0ElecstatdiscdmBefore);
                cells[y, 29].PutValue(aml.T0ElecstatdiscdmAfter);
                cells[y, 30].PutValue(aml.T0ElecstatdiscdmChanged);
                cells[y, 31].PutValue(aml.T0ElecstatdishbmBefore);
                cells[y, 32].PutValue(aml.T0ElecstatdishbmAfter);
                cells[y, 33].PutValue(aml.T0ElecstatdishbmChanged);
                cells[y, 34].PutValue(aml.T0ElecstatdismmBefore);
                cells[y, 35].PutValue(aml.T0ElecstatdismmAfter);
                cells[y, 36].PutValue(aml.T0ElecstatdismmChanged);
            }

            string filePath = @"\\suznt004\Backup\Jesse\FrontAll\FCT KANBAN\files\AML_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_") + DateTime.Now.Millisecond.ToString() + new Random().Next(1, 10000).ToString() + ".xlsx";
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

        public class AMLList
        {
            public List<AmlDiffPe> list { get; set; }
        }
    }
}
