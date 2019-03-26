using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;


namespace ApiForFOL.Controllers
{
    public class FOL_OutputController : ApiController
    {
        FOLDataClassesDataContext dc = new FOLDataClassesDataContext();

        [System.Web.Http.HttpGet]
        public RespAll getAllDetail(string version)
        {
            RespAll resp = new RespAll();
            resp.PriorForcast = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_PriorFsct] where version='" + version + "'");
            resp.CurrentForcast = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_CurFcst] where version='" + version + "'");
            resp.Actual = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_Actuals] where version='" + version + "'");
            return resp;
        }

        [System.Web.Http.HttpGet]
        public RespAll getOutputData(string version, string project, string fileType)
        {
            RespAll resp = new RespAll();
            try
            {
                resp.Code = "200";
                resp.PriorForcast = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_Actuals] where version='" + version + "' and Project='" + project + "' and fileType='PriorFcst'");
                resp.CurrentForcast = SqlHelper.ExecuteDataTable("select * from  [dbo].[FOL_Input_1_1] where version='" + version + "' and CustomerOutputCode like '%" + project + "%'");
                resp.Actual = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_Actuals] where version='" + version + "' and Project='" + project + "'  and fileType='Actual'");
                resp.OutputModules = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputModulesByCustomer]");
                //resp.ProjectMapping = null;
                resp.Message = "Success";
            }
            catch (Exception ex)
            {
                resp.Code = "10001";
                resp.Message = ex.Message;
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getProjectList()
        {
            Resp resp = new Resp();
            try
            {
                resp.Data = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_ProjectMapping]");
                resp.Code = "200";
                resp.Message = "success";
            }
            catch (Exception ex)
            {
                resp.Code = "10000";
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getActuals(string version)
        {
            Resp resp = new Resp();
            try
            {
                resp.Code = "200";
                resp.Data = SqlHelper.ExecuteDataTable("select * from [dbo].[FOL_OutputByProject_Actuals] where version = '" + version + "'");
                resp.Message = "Success";
            }
            catch (Exception ex)
            {
                resp.Code = "10000";
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp uploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            string version = httpRequest.Form["version"];
            string fileType = httpRequest.Form["fileType"];
            string project = httpRequest.Form["project"];
            HttpPostedFile file = httpRequest.Files[0];
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);
            Workbook workbook = new Workbook(filePath);
            string result = UploadOutputFile(fileType, version, project, workbook);
            Resp resp = new Resp();
            try
            {
                resp.Code = "200";
                resp.Data = null;
                resp.Message = result;
            }
            catch (Exception ex)
            {
                resp.Code = "10001";
                resp.Data = null;
                resp.Message = ex.Message;
            }
            return resp;
        }
        [System.Web.Http.HttpGet]
        public Resp getResult(string version, string project)
        {
            Resp resp = new Resp();
            string sql = "select * from [dbo].[FOL_Output_Result] where version='" + version + "' and project ='" + project + "'";
            DataTable dt = SqlHelper.ExecuteDataTable(sql);

            resp.Data = dt;
            resp.Code = "200";
            resp.Message = "";
            return resp;
        }
        [System.Web.Http.HttpPost]
        public Resp saveOutputResult(Result model)
        {

            string sql = "select * from [dbo].[FOL_Output_Result] where version='" + model.version + "' and project ='" + model.project + "'";
            DataTable dt = SqlHelper.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                sql = @"update [dbo].[FOL_Output_Result] set Value='" + model.result + "', Col = '" + model.column + "' where version='" + model.version + "' and project='" + model.project + "'";
            }
            else
            {
                sql = @"insert into [dbo].[FOL_Output_Result] values('" + model.version + "','" + model.column + "','" + model.result + "','" + model.project + "','" + DateTime.Now.ToString() + "','')";
            }
            int msg = SqlHelper.ExecuteNonQuery(sql);
            Resp resp = new Resp();
            if (msg == 0)
            {
                resp.Message = "fail";
                resp.Code = "10001";
                resp.Data = null;
            }
            else
            {
                resp.Message = "Success";
                resp.Code = "200";
                resp.Data = null;
            }
            return resp;
        }

        public string UploadOutputFile(string fileType, string version, string project, Workbook workbook)
        {
            Cells cells = workbook.Worksheets[0].Cells;
            int max = cells.MaxDataRow;
            int columns = cells.MaxColumn;
            string json;
            try
            {
                List<Item> list = new List<Item>();
                for (int i = 2; i < cells.Rows.Count; i++)
                {
                    Item item = new Item();
                    item.BPCCode = cells[i, 0].Value == null ? "" : cells[i, 0].Value.ToString();
                    List<Detail> list_details = new List<Detail>();
                    for (int j = 2; j < cells.Columns.Count; j++)
                    {
                        Detail detail = new Detail();
                        detail.month = cells[0, j].Value.ToString() + "," + cells[1, j].Value.ToString();
                        detail.money = cells[i, j].Value == null ? 0 : Convert.ToDouble(cells[i, j].Value.ToString());
                        list_details.Add(detail);
                    }
                    item.Details = list_details;
                    list.Add(item);
                }

                json = Newtonsoft.Json.JsonConvert.SerializeObject(list); //data为数据源，datatable或者list等都可以
                FOL_OutputByProject_Actuals outputItem = new FOL_OutputByProject_Actuals();
                outputItem.Version = version;
                outputItem.InsertUser = "";
                outputItem.InsertTime = DateTime.Now.ToString();
                outputItem.Value = json;
                outputItem.Project = project;
                outputItem.FileType = fileType;
                if (dc.FOL_OutputByProject_Actuals.FirstOrDefault(x => x.Version == version && x.Project == project && x.FileType == fileType) == null)
                {
                    dc.FOL_OutputByProject_Actuals.InsertOnSubmit(outputItem);
                }
                else
                {
                    dc.FOL_OutputByProject_Actuals.DeleteOnSubmit(dc.FOL_OutputByProject_Actuals.FirstOrDefault(x => x.Version == version && x.Project == project && x.FileType == fileType));
                    dc.FOL_OutputByProject_Actuals.InsertOnSubmit(outputItem);
                }
                dc.SubmitChanges();

                return "Success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }
    }
}
public class Item
{
    public string BPCCode;
    public List<Detail> Details;
}

public class Detail
{
    public string month;
    public double money;
}
public class Resp
{
    public string Code { get; set; }
    public DataTable Data { get; set; }
    public string Message { get; set; }
}

public class RespAll
{
    public string Code { get; set; }
    public DataTable PriorForcast { get; set; }
    public DataTable CurrentForcast { get; set; }
    public DataTable Actual { get; set; }
    public DataTable OutputModules { get; set; }
    //public DataTable ProjectMapping { get; set; }
    public string Message { get; set; }
}

public class Result
{
    public string version { get; set; }
    public string project { get; set; }
    public string column { get; set; }
    public string result { get; set; }
}