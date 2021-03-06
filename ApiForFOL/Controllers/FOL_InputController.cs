﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
namespace ApiForFOL.Controllers
{

    public class FOL_InputController : ApiController
    {
        FOLDataClassesDataContext dc = new FOLDataClassesDataContext();

        [System.Web.Http.HttpGet]
        public RespAll getTypedList(string version, string type, string cc)
        {
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            if (type == "6.0 Total MOH($)" || type == "7.0 Total SG&A($)")
            {
                dt1 = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_6_0_TotalMOH] where version= '" + version + " ' and type like '%" + type + "%' and Department ='" + cc + "'");
            }
            else
            {
                dt1 = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_1_1] where version= '" + version + " ' and type like '%" + type + "%'");
                dt2 = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_2_1] where version= '" + version + " ' and type like '%" + type + "%'");
                dt3 = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[FOL_Input_3_1] where version= '" + version + " ' and type like '%" + type + "%'");

            }
            RespAll resp = new RespAll();
            resp.Code = "200";
            resp.DataForecast = dt1;
            resp.DataPercent = dt2;
            resp.DataAmount = dt3;
            resp.Message = "查询成功!";
            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp getItemById(int id, string type, string menuType)
        {
            string tableName = "";
            switch (type)
            {
                case "forecast":
                    if (menuType == "6.0 Total MOH($)" || menuType == "7.0 Total SG&A($)")
                        tableName = "FOL_Input_6_0_TotalMOH";
                    else
                        tableName = "FOL_Input_1_1";
                    break;
                case "percent":
                    tableName = "FOL_Input_2_1";
                    break;
                case "amount":
                    tableName = "FOL_Input_3_1";
                    break;
                default:
                    tableName = "";
                    break;
            }
            DataTable dt = new DataTable();
            dt = SqlHelper.ExecuteDataTable("select * from [test1].[dbo].[" + tableName + "] where id = " + id);
            Resp resp = new Resp();
            resp.Code = "200";
            resp.Data = dt;
            resp.Message = "查询成功";
            return resp;
        }

        [System.Web.Http.HttpPost]
        public Resp updateItemById(Input model)
        {
            int index;
            string tableName = "";
            switch (model.DataType)
            {
                case "forecast":
                    if (model.MenuType == "6.0 Total MOH($)" || model.MenuType == "7.0 Total SG&A($)")
                        tableName = "FOL_Input_6_0_TotalMOH";
                    else
                        tableName = "FOL_Input_1_1";
                    break;
                case "percent":
                    tableName = "FOL_Input_2_1";
                    break;
                case "amount":
                    tableName = "FOL_Input_3_1";
                    break;
                default:
                    tableName = "";
                    break;
            }
            index = SqlHelper.ExecuteNonQuery("update " + tableName + " set " +
            "Period1 = " + model.Period1 + ", " +
            "Period2 = " + model.Period2 + ", " +
            "Period3 = " + model.Period3 + ", " +
            "Period4 = " + model.Period4 + ", " +
            "Period5 = " + model.Period5 + ", " +
            "Period6 = " + model.Period6 + ", " +
            "Period7 = " + model.Period7 + ", " +
            "Period8 = " + model.Period8 + ", " +
            "Period9 = " + model.Period9 + ", " +
            "Period10 = " + model.Period10 + ", " +
            "Period11 = " + model.Period11 + ", " +
            "Period12 = " + model.Period12 + ", " +
            "Period13 = " + model.Period13 + ", " +
            "Period14 = " + model.Period14 + ", " +
            "Period15 = " + model.Period15 + ", " +
            "Period16 = " + model.Period16 +
            "where id = " + model.Id);
            Resp resp = new Resp();
            if (index != 0)
            {
                resp.Code = "200";
                resp.Data = null;
                resp.Message = "修改成功";
            }
            else
            {
                resp.Code = "10001";
                resp.Data = null;
                resp.Message = "修改失败";
            }

            return resp;
        }

        [System.Web.Http.HttpGet]
        public Resp calculateByType(string type, string version, string cc)
        {
            FOL_InputFormulas fi = new FOL_InputFormulas();
            Resp resp = new Resp();
            string message = "";
            if (type == "3.1 PPV(%&$)" || type == "4.2 Freight in(%)" || type == "4.3 Freight out(%)" || type == "4.4 EDM(%)" || type == "4.5 MCOSO(%)" || type == "4.6 Subcontract($)" || type == "4.7 Inv. Reserve($)")
            {
                message = fi.Calculate3_1(version, type);
            }
            else if (type == "6.2 Depn(%)" || type == "6.3 FUS.(%)" || type == "6.4 MOHO(%)" || type == "6.5 OEUP(%)" || type == "6.6 ELEC(%)")
            {
                message = fi.Calculate6_2(version, type);
            }
            else if (type == "7.1 G.A.(%)" || type == "7.2 S.M.(%)")
            {
                message = fi.Calculate7_1(version, type, cc);
            }
            else
            {
                switch (type)
                {
                    case "1.0 Total Sales($)":
                        message = fi.Calculate1_0(version, type);
                        break;
                    case "2.0 MCOS($)":
                        message = fi.Calculate2_0(version, type);
                        break;
                    case "2.1 Std VAM % from Ops(%)":
                        message = fi.Calculate2_1(version, type);
                        break;
                    case "2.4 COGS Reversal - OT Contract":
                        message = fi.Calculate2_4(version, type);
                        break;
                    case "3.0 Material Margin($)":
                        message = fi.Calculate3_0(version, type);
                        break;
                    case "4.1 Material Loss(%)":
                        message = fi.Calculate4_1(version, type);
                        break;
                    case "5.1 DL(%)":
                        message = fi.Calculate5_1(version, type);
                        break;
                    case "6.0 Total MOH($)":
                        message = fi.Calculate6_0(version, type, cc);
                        break;
                    case "6.1 IDL(%&#)":
                        message = fi.Calculate6_1(version, type, cc);
                        break;
                    //case "8.0 Corp. Alloc.(%)":
                    //    message = fi.Calculate8_0(version, type);
                    //    break;
                    default:
                        message = "Default";
                        break;
                }

            }
            resp.Code = "200";
            resp.Message = message;
            return resp;

        }

        /// <summary>
        /// 接收前端上传单个文件
        /// </summary>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        public Resp uploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            string type = httpRequest.Form["type"];
            string version = httpRequest.Form["version"];
            string fileType = httpRequest.Form["fileType"];
            string CC = httpRequest.Form["cc"];
            HttpPostedFile file = httpRequest.Files[0];
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);
            Workbook workbook = new Workbook(filePath);
            string result;
            switch (fileType)
            {
                case "1":
                    if (type == "6.0 Total MOH($)" || type == "7.0 Total SG&A($)")
                        result = Upload6_0(type, version, workbook, CC);
                    else
                        result = Upload1_1(type, version, workbook, CC);
                    break;
                case "2":
                    result = Upload2_1(type, version, workbook, CC);
                    break;
                case "3":
                    result = Upload3_1(type, version, workbook, CC);
                    break;
                default:
                    result = "";
                    break;
            }
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

        public class RespAll
        {
            public string Code { get; set; }
            public DataTable DataForecast { get; set; }
            public DataTable DataPercent { get; set; }
            public DataTable DataAmount { get; set; }
            public string Message { get; set; }
        }

        public string Upload1_1(string insertType, string version, Workbook workbook, string CC)
        {
            Cells cells = workbook.Worksheets[0].Cells;
            int max = cells.MaxDataRow;
            int columns = cells.MaxColumn;
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            try
            {
                for (int i = 1; i <= max; i++)
                {
                    FOL_Input_1_1 item = new FOL_Input_1_1
                    {
                        GLBPCCode = cells[i, 0].Value == null ? "" : cells[i, 0].Value.ToString(),
                        GLBPCDescription = cells[i, 1].Value == null ? "" : cells[i, 1].Value.ToString(),
                        GLOutputCode = cells[i, 2].Value == null ? "" : cells[i, 2].Value.ToString(),
                        CustomerBPCCode = cells[i, 3].Value == null ? "" : cells[i, 3].Value.ToString(),
                        DIM1 = cells[i, 4].Value == null ? "" : cells[i, 4].Value.ToString(),
                        Order = cells[i, 5].Value == null ? "" : cells[i, 5].Value.ToString(),
                        CustomerOutputCode = cells[i, 6].Value == null ? "" : cells[i, 6].Value.ToString(),
                        BU = cells[i, 7].Value == null ? "" : cells[i, 7].Value.ToString(),
                        Segment = cells[i, 8].Value == null ? "" : cells[i, 8].Value.ToString(),

                        Period1 = Convert.ToDouble(cells[i, 9].Value ?? 0),
                        Period2 = Convert.ToDouble(cells[i, 10].Value ?? 0),
                        Period3 = Convert.ToDouble(cells[i, 11].Value ?? 0),
                        Period4 = Convert.ToDouble(cells[i, 12].Value ?? 0),
                        Period5 = Convert.ToDouble(cells[i, 13].Value ?? 0),
                        Period6 = Convert.ToDouble(cells[i, 14].Value ?? 0),
                        Period7 = Convert.ToDouble(cells[i, 15].Value ?? 0),
                        Period8 = Convert.ToDouble(cells[i, 16].Value ?? 0),
                        Period9 = Convert.ToDouble(cells[i, 17].Value ?? 0),
                        Period10 = Convert.ToDouble(cells[i, 18].Value ?? 0),
                        Period11 = Convert.ToDouble(cells[i, 19].Value ?? 0),
                        Period12 = Convert.ToDouble(cells[i, 20].Value ?? 0),
                        Period13 = Convert.ToDouble(cells[i, 21].Value ?? 0),
                        Period14 = Convert.ToDouble(cells[i, 22].Value ?? 0),
                        Period15 = Convert.ToDouble(cells[i, 23].Value ?? 0),
                        Period16 = Convert.ToDouble(cells[i, 24].Value ?? 0),

                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                        Type = insertType,
                        Department = CC,
                    };
                    list.Add(item);

                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == insertType && x.Department == CC).ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                return "Success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public string Upload2_1(string insertType, string version, Workbook workbook, string CC)
        {
            Cells cells = workbook.Worksheets[0].Cells;
            int max = cells.MaxDataRow;
            int columns = cells.MaxColumn;
            List<FOL_Input_2_1> list = new List<FOL_Input_2_1>();
            try
            {
                for (int i = 1; i <= cells.MaxDataRow; i++)
                {
                    FOL_Input_2_1 item = new FOL_Input_2_1
                    {
                        GLBPCCode = cells[i, 0].Value == null ? "" : cells[i, 0].Value.ToString(),
                        GLBPCDescription = cells[i, 1].Value == null ? "" : cells[i, 1].Value.ToString(),
                        GLOutputCode = cells[i, 2].Value == null ? "" : cells[i, 2].Value.ToString(),
                        CustomerBPCCode = cells[i, 3].Value == null ? "" : cells[i, 3].Value.ToString(),
                        DIM1 = cells[i, 4].Value == null ? "" : cells[i, 4].Value.ToString(),
                        Order = cells[i, 5].Value == null ? "" : cells[i, 5].Value.ToString(),
                        CustomerOutputCode = cells[i, 6].Value == null ? "" : cells[i, 6].Value.ToString(),
                        BU = cells[i, 7].Value == null ? "" : cells[i, 7].Value.ToString(),
                        Segment = cells[i, 8].Value == null ? "" : cells[i, 8].Value.ToString(),

                        Period1 = Convert.ToDouble(cells[i, 9].Value ?? 0),
                        Period2 = Convert.ToDouble(cells[i, 10].Value ?? 0),
                        Period3 = Convert.ToDouble(cells[i, 11].Value ?? 0),
                        Period4 = Convert.ToDouble(cells[i, 12].Value ?? 0),
                        Period5 = Convert.ToDouble(cells[i, 13].Value ?? 0),
                        Period6 = Convert.ToDouble(cells[i, 14].Value ?? 0),
                        Period7 = Convert.ToDouble(cells[i, 15].Value ?? 0),
                        Period8 = Convert.ToDouble(cells[i, 16].Value ?? 0),
                        Period9 = Convert.ToDouble(cells[i, 17].Value ?? 0),
                        Period10 = Convert.ToDouble(cells[i, 18].Value ?? 0),
                        Period11 = Convert.ToDouble(cells[i, 19].Value ?? 0),
                        Period12 = Convert.ToDouble(cells[i, 20].Value ?? 0),
                        Period13 = Convert.ToDouble(cells[i, 21].Value ?? 0),
                        Period14 = Convert.ToDouble(cells[i, 22].Value ?? 0),
                        Period15 = Convert.ToDouble(cells[i, 23].Value ?? 0),
                        Period16 = Convert.ToDouble(cells[i, 24].Value ?? 0),

                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                        Type = insertType,
                        Department = CC,
                    };
                    list.Add(item);
                }

                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Version == version && x.Type == insertType && x.Department == CC).ToList());
                dc.FOL_Input_2_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                return "Success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public string Upload3_1(string insertType, string version, Workbook workbook, string CC)
        {
            Cells cells = workbook.Worksheets[0].Cells;
            int max = cells.MaxDataRow;
            int columns = cells.MaxColumn;
            List<FOL_Input_3_1> list = new List<FOL_Input_3_1>();
            try
            {
                for (int i = 1; i <= cells.MaxDataRow; i++)
                {
                    FOL_Input_3_1 item = new FOL_Input_3_1
                    {
                        GLBPCCode = cells[i, 0].Value == null ? "" : cells[i, 0].Value.ToString(),
                        GLBPCDescription = cells[i, 1].Value == null ? "" : cells[i, 1].Value.ToString(),
                        GLOutputCode = cells[i, 2].Value == null ? "" : cells[i, 2].Value.ToString(),
                        CustomerBPCCode = cells[i, 3].Value == null ? "" : cells[i, 3].Value.ToString(),
                        DIM1 = cells[i, 4].Value == null ? "" : cells[i, 4].Value.ToString(),
                        Order = cells[i, 5].Value == null ? "" : cells[i, 5].Value.ToString(),
                        CustomerOutputCode = cells[i, 6].Value == null ? "" : cells[i, 6].Value.ToString(),
                        BU = cells[i, 7].Value == null ? "" : cells[i, 7].Value.ToString(),
                        Segment = cells[i, 8].Value == null ? "" : cells[i, 8].Value.ToString(),

                        Period1 = Convert.ToDouble(cells[i, 9].Value ?? 0),
                        Period2 = Convert.ToDouble(cells[i, 10].Value ?? 0),
                        Period3 = Convert.ToDouble(cells[i, 11].Value ?? 0),
                        Period4 = Convert.ToDouble(cells[i, 12].Value ?? 0),
                        Period5 = Convert.ToDouble(cells[i, 13].Value ?? 0),
                        Period6 = Convert.ToDouble(cells[i, 14].Value ?? 0),
                        Period7 = Convert.ToDouble(cells[i, 15].Value ?? 0),
                        Period8 = Convert.ToDouble(cells[i, 16].Value ?? 0),
                        Period9 = Convert.ToDouble(cells[i, 17].Value ?? 0),
                        Period10 = Convert.ToDouble(cells[i, 18].Value ?? 0),
                        Period11 = Convert.ToDouble(cells[i, 19].Value ?? 0),
                        Period12 = Convert.ToDouble(cells[i, 20].Value ?? 0),
                        Period13 = Convert.ToDouble(cells[i, 21].Value ?? 0),
                        Period14 = Convert.ToDouble(cells[i, 22].Value ?? 0),
                        Period15 = Convert.ToDouble(cells[i, 23].Value ?? 0),
                        Period16 = Convert.ToDouble(cells[i, 24].Value ?? 0),

                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                        Type = insertType,
                        Department = CC,
                    };
                    list.Add(item);
                }

                dc.FOL_Input_3_1.DeleteAllOnSubmit(dc.FOL_Input_3_1.Where(x => x.Version == version && x.Type == insertType && x.Department == CC).ToList());
                dc.FOL_Input_3_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                return "Success";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        public string Upload6_0(string insertType, string version, Workbook workbook, string CC)
        {
            Cells cells = workbook.Worksheets[0].Cells;
            int max = cells.MaxDataRow;
            int columns = cells.MaxColumn;
            List<FOL_Input_6_0_TotalMOH> list = new List<FOL_Input_6_0_TotalMOH>();
            try
            {
                for (int i = 1; i <= cells.MaxDataRow; i++)
                {
                    FOL_Input_6_0_TotalMOH item = new FOL_Input_6_0_TotalMOH
                    {
                        GLBPCCode = cells[i, 0].Value == null ? "" : cells[i, 0].Value.ToString(),
                        GLBPCDescription = cells[i, 1].Value == null ? "" : cells[i, 1].Value.ToString(),
                        GLOutputCode = cells[i, 2].Value == null ? "" : cells[i, 2].Value.ToString(),
                        PARENTH1 = cells[i, 3].Value == null ? "" : cells[i, 3].Value.ToString(),
                        UploadCode = cells[i, 4].Value == null ? "" : cells[i, 4].Value.ToString(),
                        BPCOutputCode = cells[i, 5].Value == null ? "" : cells[i, 5].Value.ToString(),
                        BU = cells[i, 6].Value == null ? "" : cells[i, 6].Value.ToString(),
                        Segment = cells[i, 7].Value == null ? "" : cells[i, 7].Value.ToString(),
                        Department = CC,
                        IsSubtotal = dc.FOL_ExpenseMappingManagement_.Where(x => x.GLBPCCode.ToLower() == cells[i, 0].Value.ToString().ToLower().Trim()).FirstOrDefault().Formulas ?? "",

                        Period1 = Convert.ToDouble(cells[i, 8].Value ?? 0),
                        Period2 = Convert.ToDouble(cells[i, 9].Value ?? 0),
                        Period3 = Convert.ToDouble(cells[i, 10].Value ?? 0),
                        Period4 = Convert.ToDouble(cells[i, 11].Value ?? 0),
                        Period5 = Convert.ToDouble(cells[i, 12].Value ?? 0),
                        Period6 = Convert.ToDouble(cells[i, 13].Value ?? 0),
                        Period7 = Convert.ToDouble(cells[i, 14].Value ?? 0),
                        Period8 = Convert.ToDouble(cells[i, 15].Value ?? 0),
                        Period9 = Convert.ToDouble(cells[i, 16].Value ?? 0),
                        Period10 = Convert.ToDouble(cells[i, 17].Value ?? 0),
                        Period11 = Convert.ToDouble(cells[i, 18].Value ?? 0),
                        Period12 = Convert.ToDouble(cells[i, 19].Value ?? 0),
                        Period13 = Convert.ToDouble(cells[i, 20].Value ?? 0),
                        Period14 = Convert.ToDouble(cells[i, 21].Value ?? 0),
                        Period15 = Convert.ToDouble(cells[i, 22].Value ?? 0),
                        Period16 = Convert.ToDouble(cells[i, 23].Value ?? 0),

                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                        Type = insertType,
                    };
                    list.Add(item);
                }

                dc.FOL_Input_6_0_TotalMOH.DeleteAllOnSubmit(dc.FOL_Input_6_0_TotalMOH.Where(x => x.Version == version && x.Type == insertType && x.Department == CC).ToList());
                dc.FOL_Input_6_0_TotalMOH.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                return "Success";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }
    }

    public class Resp
    {
        public string Code { get; set; }
        public DataTable Data { get; set; }
        public string Message { get; set; }
    }

    public class Input
    {
        public int Id { get; set; }
        public string DataType { get; set; }
        public string MenuType { get; set; }
        public float Period1 { get; set; }
        public float Period2 { get; set; }
        public float Period3 { get; set; }
        public float Period4 { get; set; }
        public float Period5 { get; set; }
        public float Period6 { get; set; }
        public float Period7 { get; set; }
        public float Period8 { get; set; }
        public float Period9 { get; set; }
        public float Period10 { get; set; }
        public float Period11 { get; set; }
        public float Period12 { get; set; }
        public float Period13 { get; set; }
        public float Period14 { get; set; }
        public float Period15 { get; set; }
        public float Period16 { get; set; }
    }
}
