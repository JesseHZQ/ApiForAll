using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForFOL
{
    public class FOL_InputFormulas
    {
        FOLDataClassesDataContext dc = new FOLDataClassesDataContext();
        public FOL_InputFormulas()
        {

        }

        //1.0
        public string Calculate1_0(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list4 = dc.FOL_Input_1_1.Where(x => x.Type == "1.4 Rev Recog - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_1_1> list5 = dc.FOL_Input_1_1.Where(x => x.Type == "1.5 Rev Reversal - OT Contract" && x.Version == version).ToList();
            if (list1.Count == 0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list4.Count == 0)
            { message = message + "该月份的1.4 Rev Recog - OT Contract数据为空;"; return message; }
            if (list5.Count == 0)
            { message = message + "该月份的1.5 Rev Reversal - OT Contract数据为空;"; return message; }
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            try
            {
                foreach (var item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = "30_SALES",
                        GLBPCDescription = "Total Sales",
                        GLOutputCode = "30_SALES" + " - " + item.DIM1.Substring(0, 3),
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = item.Period1 + list2[i].Period1 + list3[i].Period1 + (list4[i].Period1 == null ? 0 : list4[i].Period1) + (list5[i].Period1 ?? 0),
                        Period2 = item.Period2 + list2[i].Period2 + list3[i].Period2 + (list4[i].Period2 == null ? 0 : list4[i].Period2) + (list5[i].Period2 ?? 0),
                        Period3 = item.Period3 + list2[i].Period3 + list3[i].Period3 + (list4[i].Period3 == null ? 0 : list4[i].Period3) + (list5[i].Period3 ?? 0),
                        Period4 = item.Period4 + list2[i].Period4 + list3[i].Period4 + (list4[i].Period4 == null ? 0 : list4[i].Period4) + (list5[i].Period4 ?? 0),
                        Period5 = item.Period5 + list2[i].Period5 + list3[i].Period5 + (list4[i].Period5 == null ? 0 : list4[i].Period5) + (list5[i].Period5 ?? 0),
                        Period6 = item.Period6 + list2[i].Period6 + list3[i].Period6 + (list4[i].Period6 == null ? 0 : list4[i].Period6) + (list5[i].Period6 ?? 0),
                        Period7 = item.Period7 + list2[i].Period7 + list3[i].Period7 + (list4[i].Period7 == null ? 0 : list4[i].Period7) + (list5[i].Period7 ?? 0),
                        Period8 = item.Period8 + list2[i].Period8 + list3[i].Period8 + (list4[i].Period8 == null ? 0 : list4[i].Period8) + (list5[i].Period8 ?? 0),
                        Period9 = item.Period9 + list2[i].Period9 + list3[i].Period9 + (list4[i].Period9 == null ? 0 : list4[i].Period9) + (list5[i].Period9 ?? 0),
                        Period10 = item.Period10 + list2[i].Period10 + list3[i].Period10 + (list4[i].Period10 == null ? 0 : list4[i].Period10) + (list5[i].Period10 ?? 0),
                        Period11 = item.Period11 + list2[i].Period11 + list3[i].Period11 + (list4[i].Period11 == null ? 0 : list4[i].Period11) + (list5[i].Period11 ?? 0),
                        Period12 = item.Period12 + list2[i].Period12 + list3[i].Period12 + (list4[i].Period12 == null ? 0 : list4[i].Period12) + (list5[i].Period12 ?? 0),
                        Period13 = item.Period13 + list2[i].Period13 + list3[i].Period13 + (list4[i].Period13 == null ? 0 : list4[i].Period13) + (list5[i].Period13 ?? 0),
                        Period14 = item.Period14 + list2[i].Period14 + list3[i].Period14 + (list4[i].Period14 == null ? 0 : list4[i].Period14) + (list5[i].Period14 ?? 0),
                        Period15 = item.Period15 + list2[i].Period15 + list3[i].Period15 + (list4[i].Period15 == null ? 0 : list4[i].Period15) + (list5[i].Period15 ?? 0),
                        Period16 = item.Period16 + list2[i].Period16 + list3[i].Period16 + (list4[i].Period16 == null ? 0 : list4[i].Period16) + (list5[i].Period16 ?? 0),
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }

                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "1.0 Total Sales($)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }

        //2.0
        public string Calculate2_0(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "2.2 MCOS Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list3 = dc.FOL_Input_1_1.Where(x => x.Type == "2.3 COGS Recog - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_1_1> list4 = dc.FOL_Input_1_1.Where(x => x.Type == "2.4 COGS Reversal - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_1_1> list_TotalSales = dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "1.0 Total Sales($)").ToList();
            if (list1.Count == 0)
            { message = message + "该月份的2.1 Std VAM % from Ops(%)数据为空"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的2.2 MCOS Recoveries($)数据为空"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的2.3 COGS Recog - OT Contract数据为空"; return message; }
            if (list4.Count == 0)
            { message = message + "该月份的2.4 COGS Reversal - OT Contract数据为空"; return message; }

            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            #region
            try
            {
                foreach (var item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = "400",
                        GLBPCDescription = "Total Standard Material Cost of Sales",
                        GLOutputCode = "400" + " - " + list1[i].DIM1.Substring(0, 3),
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = item.Period1 + list2[i].Period1 + list3[i].Period1 + (list4[i].Period1 == null ? 0 : list4[i].Period1),
                        Period2 = item.Period2 + list2[i].Period2 + list3[i].Period2 + (list4[i].Period2 == null ? 0 : list4[i].Period2),
                        Period3 = item.Period3 + list2[i].Period3 + list3[i].Period3 + (list4[i].Period3 == null ? 0 : list4[i].Period3),
                        Period4 = item.Period4 + list2[i].Period4 + list3[i].Period4 + (list4[i].Period4 == null ? 0 : list4[i].Period4),
                        Period5 = item.Period5 + list2[i].Period5 + list3[i].Period5 + (list4[i].Period5 == null ? 0 : list4[i].Period5),
                        Period6 = item.Period6 + list2[i].Period6 + list3[i].Period6 + (list4[i].Period6 == null ? 0 : list4[i].Period6),
                        Period7 = item.Period7 + list2[i].Period7 + list3[i].Period7 + (list4[i].Period7 == null ? 0 : list4[i].Period7),
                        Period8 = item.Period8 + list2[i].Period8 + list3[i].Period8 + (list4[i].Period8 == null ? 0 : list4[i].Period8),
                        Period9 = item.Period9 + list2[i].Period9 + list3[i].Period9 + (list4[i].Period9 == null ? 0 : list4[i].Period9),
                        Period10 = item.Period10 + list2[i].Period10 + list3[i].Period10 + (list4[i].Period10 == null ? 0 : list4[i].Period10),
                        Period11 = item.Period11 + list2[i].Period11 + list3[i].Period11 + (list4[i].Period11 == null ? 0 : list4[i].Period11),
                        Period12 = item.Period12 + list2[i].Period12 + list3[i].Period12 + (list4[i].Period12 == null ? 0 : list4[i].Period12),
                        Period13 = item.Period13 + list2[i].Period13 + list3[i].Period13 + (list4[i].Period13 == null ? 0 : list4[i].Period13),
                        Period14 = item.Period14 + list2[i].Period14 + list3[i].Period14 + (list4[i].Period14 == null ? 0 : list4[i].Period14),
                        Period15 = item.Period15 + list2[i].Period15 + list3[i].Period15 + (list4[i].Period15 == null ? 0 : list4[i].Period15),
                        Period16 = item.Period16 + list2[i].Period16 + list3[i].Period16 + (list4[i].Period16 == null ? 0 : list4[i].Period16),
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "2.0 MCOS($)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
                return message;
            }
            #endregion

            #region
            if (list_TotalSales.Count == 0)
            { message = message + "该月份的1.0 Total Sales($)数据为空"; return message; }
            List<FOL_Input_2_1> list2_1 = new List<FOL_Input_2_1>();
            try
            {
                i = 0;
                foreach (var item in list)
                {
                    FOL_Input_2_1 newItem = new FOL_Input_2_1
                    {
                        GLBPCCode = item.GLBPCCode,
                        GLBPCDescription = item.GLBPCDescription,
                        GLOutputCode = item.GLOutputCode,
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = (list_TotalSales[i].Period1 == 0) ? 0 : (1 - (item.Period1 / list_TotalSales[i].Period1)),
                        Period2 = (list_TotalSales[i].Period2 == 0) ? 0 : (1 - (item.Period2 / list_TotalSales[i].Period2)),
                        Period3 = (list_TotalSales[i].Period3 == 0) ? 0 : (1 - (item.Period3 / list_TotalSales[i].Period3)),
                        Period4 = (list_TotalSales[i].Period4 == 0) ? 0 : (1 - (item.Period4 / list_TotalSales[i].Period4)),
                        Period5 = (list_TotalSales[i].Period5 == 0) ? 0 : (1 - (item.Period5 / list_TotalSales[i].Period5)),
                        Period6 = (list_TotalSales[i].Period6 == 0) ? 0 : (1 - (item.Period6 / list_TotalSales[i].Period6)),
                        Period7 = (list_TotalSales[i].Period7 == 0) ? 0 : (1 - (item.Period7 / list_TotalSales[i].Period7)),
                        Period8 = (list_TotalSales[i].Period8 == 0) ? 0 : (1 - (item.Period8 / list_TotalSales[i].Period8)),
                        Period9 = (list_TotalSales[i].Period9 == 0) ? 0 : (1 - (item.Period9 / list_TotalSales[i].Period9)),
                        Period10 = (list_TotalSales[i].Period10 == 0) ? 0 : (1 - (item.Period10 / list_TotalSales[i].Period10)),
                        Period11 = (list_TotalSales[i].Period11 == 0) ? 0 : (1 - (item.Period11 / list_TotalSales[i].Period11)),
                        Period12 = (list_TotalSales[i].Period12 == 0) ? 0 : (1 - (item.Period12 / list_TotalSales[i].Period12)),
                        Period13 = (list_TotalSales[i].Period13 == 0) ? 0 : (1 - (item.Period13 / list_TotalSales[i].Period13)),
                        Period14 = (list_TotalSales[i].Period14 == 0) ? 0 : (1 - (item.Period14 / list_TotalSales[i].Period14)),
                        Period15 = (list_TotalSales[i].Period15 == 0) ? 0 : (1 - (item.Period15 / list_TotalSales[i].Period15)),
                        Period16 = (list_TotalSales[i].Period16 == 0) ? 0 : (1 - (item.Period16 / list_TotalSales[i].Period16)),

                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list2_1.Add(newItem);
                }
                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Version == version && x.Type == "2.0 MCOS($)").ToList());
                dc.FOL_Input_2_1.InsertAllOnSubmit(list2_1);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            #endregion
            return message;
        }

        //2.1
        public string Calculate2_1(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_2_1> list3 = dc.FOL_Input_2_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            if (list1.Count == 0)
            { message = message + "该月份的1.1 Gross Sales - External($)数据为空"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的2.1 Std VAM % from Ops(%)的百分比数据为空"; return message; }

            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            try
            {
                foreach (var item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = list3[i].GLBPCCode,
                        GLBPCDescription = list3[i].GLBPCDescription,
                        GLOutputCode = list3[i].GLOutputCode,
                        CustomerBPCCode = list3[i].CustomerBPCCode,
                        DIM1 = list3[i].DIM1,
                        Order = list3[i].Order,
                        CustomerOutputCode = list3[i].CustomerOutputCode,
                        BU = list3[i].BU,
                        Segment = item.Segment,
                        Period1 = (item.Period1 + list2[i].Period1) * (1 - list3[i].Period1),
                        Period2 = (item.Period2 + list2[i].Period2) * (1 - list3[i].Period2),
                        Period3 = (item.Period3 + list2[i].Period3) * (1 - list3[i].Period3),
                        Period4 = (item.Period4 + list2[i].Period4) * (1 - list3[i].Period4),
                        Period5 = (item.Period5 + list2[i].Period5) * (1 - list3[i].Period5),
                        Period6 = (item.Period6 + list2[i].Period6) * (1 - list3[i].Period6),
                        Period7 = (item.Period7 + list2[i].Period7) * (1 - list3[i].Period7),
                        Period8 = (item.Period8 + list2[i].Period8) * (1 - list3[i].Period8),
                        Period9 = (item.Period9 + list2[i].Period9) * (1 - list3[i].Period9),
                        Period10 = (item.Period10 + list2[i].Period10) * (1 - list3[i].Period10),
                        Period11 = (item.Period11 + list2[i].Period11) * (1 - list3[i].Period11),
                        Period12 = (item.Period12 + list2[i].Period12) * (1 - list3[i].Period12),
                        Period13 = (item.Period13 + list2[i].Period13) * (1 - list3[i].Period13),
                        Period14 = (item.Period14 + list2[i].Period14) * (1 - list3[i].Period14),
                        Period15 = (item.Period15 + list2[i].Period15) * (1 - list3[i].Period15),
                        Period16 = (item.Period16 + list2[i].Period16) * (1 - list3[i].Period16),

                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "2.1 Std VAM % from Ops(%)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }

        //2.4 
        public string Calculate2_4(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "2.3 COGS Recog - OT Contract" && x.Version == version).ToList();
            if (list1.Count == 0)
            { message = "该月份的2.3 COGS Recog - OT Contract数据为空;"; return message; }
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            try
            {
                foreach (FOL_Input_1_1 item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = "4040",
                        GLBPCDescription = "COGS Reversed from OT Contracts",
                        GLOutputCode = "4040" + " - " + item.DIM1.Substring(0, 3),
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period2 = -item.Period1,
                        Period3 = -item.Period2,
                        Period4 = -item.Period3,
                        Period5 = -item.Period4,
                        Period6 = -item.Period5,
                        Period7 = -item.Period6,
                        Period8 = -item.Period7,
                        Period9 = -item.Period8,
                        Period10 = -item.Period9,
                        Period11 = -item.Period10,
                        Period12 = -item.Period11,
                        Period13 = -item.Period12,
                        Period14 = -item.Period13,
                        Period15 = -item.Period14,
                        Period16 = -item.Period15,
                        Type = type,
                        Version = item.Version,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                    };
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "2.4 COGS Reversal - OT Contract").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }

        //3.0 
        public string Calculate3_0(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "3.1 PPV(%&$)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "3.2 FCP-PPV($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list3 = dc.FOL_Input_1_1.Where(x => x.Type == "3.3 Alloc-PPV($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list_percent = dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "1.0 Total Sales($)").ToList();
            if (list1.Count == 0)
            { message = "该月份的3.1 PPV(%&$)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的3.2 FCP-PPV($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的3.3 Alloc-PPV($)数据为空;"; return message; }
            if (list_percent.Count == 0)
            { message = message + "该月份的1.0 Total Sales($)数据为空;"; return message; }
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            try
            {
                #region
                foreach (var item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = "PPV",
                        GLBPCDescription = "Material Margin",
                        GLOutputCode = "PPV " + "- " + list1[i].DIM1.Substring(0, 3),
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = (item.Period1 ?? 0) + (list2[i].Period1 ?? 0) + (list3[i].Period1 ?? 0),
                        Period2 = (item.Period2 ?? 0) + (list2[i].Period2 ?? 0) + (list3[i].Period2 ?? 0),
                        Period3 = (item.Period3 ?? 0) + (list2[i].Period3 ?? 0) + (list3[i].Period3 ?? 0),
                        Period4 = (item.Period4 ?? 0) + (list2[i].Period4 ?? 0) + (list3[i].Period4 ?? 0),
                        Period5 = (item.Period5 ?? 0) + (list2[i].Period5 ?? 0) + (list3[i].Period5 ?? 0),
                        Period6 = (item.Period6 ?? 0) + (list2[i].Period6 ?? 0) + (list3[i].Period6 ?? 0),
                        Period7 = (item.Period7 ?? 0) + (list2[i].Period7 ?? 0) + (list3[i].Period7 ?? 0),
                        Period8 = (item.Period8 ?? 0) + (list2[i].Period8 ?? 0) + (list3[i].Period8 ?? 0),
                        Period9 = (item.Period9 ?? 0) + (list2[i].Period9 ?? 0) + (list3[i].Period9 ?? 0),
                        Period10 = (item.Period10 ?? 0) + (list2[i].Period10 ?? 0) + (list3[i].Period10 ?? 0),
                        Period11 = (item.Period11 ?? 0) + (list2[i].Period11 ?? 0) + (list3[i].Period11 ?? 0),
                        Period12 = (item.Period12 ?? 0) + (list2[i].Period12 ?? 0) + (list3[i].Period12 ?? 0),
                        Period13 = (item.Period13 ?? 0) + (list2[i].Period13 ?? 0) + (list3[i].Period13 ?? 0),
                        Period14 = (item.Period14 ?? 0) + (list2[i].Period14 ?? 0) + (list3[i].Period14 ?? 0),
                        Period15 = (item.Period15 ?? 0) + (list2[i].Period15 ?? 0) + (list3[i].Period15 ?? 0),
                        Period16 = (item.Period16 ?? 0) + (list2[i].Period16 ?? 0) + (list3[i].Period16 ?? 0),
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }
                #endregion
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "3.0 Material Margin ($)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
                return message;
            }

            #region
            List<FOL_Input_2_1> list2_1 = new List<FOL_Input_2_1>();
            try
            {
                i = 0;
                foreach (var item in list)
                {
                    FOL_Input_2_1 newItem = new FOL_Input_2_1
                    {
                        GLBPCCode = item.GLBPCCode,
                        GLBPCDescription = item.GLBPCDescription,
                        GLOutputCode = item.GLOutputCode,
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = (list_percent[i].Period1 == 0) ? 0 : (item.Period1 / list_percent[i].Period1),
                        Period2 = (list_percent[i].Period2 == 0) ? 0 : (item.Period2 / list_percent[i].Period2),
                        Period3 = (list_percent[i].Period3 == 0) ? 0 : (item.Period3 / list_percent[i].Period3),
                        Period4 = (list_percent[i].Period4 == 0) ? 0 : (item.Period4 / list_percent[i].Period4),
                        Period5 = (list_percent[i].Period5 == 0) ? 0 : (item.Period5 / list_percent[i].Period5),
                        Period6 = (list_percent[i].Period6 == 0) ? 0 : (item.Period6 / list_percent[i].Period6),
                        Period7 = (list_percent[i].Period7 == 0) ? 0 : (item.Period7 / list_percent[i].Period7),
                        Period8 = (list_percent[i].Period8 == 0) ? 0 : (item.Period8 / list_percent[i].Period8),
                        Period9 = (list_percent[i].Period9 == 0) ? 0 : (item.Period9 / list_percent[i].Period9),
                        Period10 = (list_percent[i].Period10 == 0) ? 0 : (item.Period10 / list_percent[i].Period10),
                        Period11 = (list_percent[i].Period11 == 0) ? 0 : (item.Period11 / list_percent[i].Period11),
                        Period12 = (list_percent[i].Period12 == 0) ? 0 : (item.Period12 / list_percent[i].Period12),
                        Period13 = (list_percent[i].Period13 == 0) ? 0 : (item.Period13 / list_percent[i].Period13),
                        Period14 = (list_percent[i].Period14 == 0) ? 0 : (item.Period14 / list_percent[i].Period14),
                        Period15 = (list_percent[i].Period15 == 0) ? 0 : (item.Period15 / list_percent[i].Period15),
                        Period16 = (list_percent[i].Period16 == 0) ? 0 : (item.Period15 / list_percent[i].Period15),
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list2_1.Add(newItem);
                }
                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Version == version && x.Type == "3.0 Material Margin ($)").ToList());
                dc.FOL_Input_2_1.InsertAllOnSubmit(list2_1);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            #endregion

            return message;
        }

        //3.1, 4.2-4.7
        public string Calculate3_1(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = new List<FOL_Input_1_1>();
            List<FOL_Input_1_1> list2 = new List<FOL_Input_1_1>();
            List<FOL_Input_1_1> list3 = new List<FOL_Input_1_1>();
            List<FOL_Input_2_1> list_percent = new List<FOL_Input_2_1>();
            List<FOL_Input_3_1> list_amount = new List<FOL_Input_3_1>();

            list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            list3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();

            list_percent = dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList();
            list_amount = dc.FOL_Input_3_1.Where(x => x.Type == type && x.Version == version).ToList();

            if (list1.Count == 0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list_percent.Count == 0)
            { message = message + "该月份的" + type + "百分比数据为空;"; return message; }
            if (list_amount.Count == 0)
            { message = message + "该月份的" + type + "Amount数据为空;"; return message; }
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            try
            {
                foreach (var item in list1)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = list_percent[i].GLBPCCode,
                        GLBPCDescription = list_percent[i].GLBPCDescription,
                        GLOutputCode = list_percent[i].GLOutputCode,
                        CustomerBPCCode = list_percent[i].CustomerBPCCode,
                        DIM1 = list_percent[i].DIM1,
                        Order = list_percent[i].Order,
                        CustomerOutputCode = list_percent[i].CustomerOutputCode,
                        BU = list_percent[i].BU,
                        Segment = list_percent[i].Segment,
                        Period1 = (item.Period1 + list2[i].Period1 + list3[i].Period1) * list_percent[i].Period1 + list_amount[i].Period1,
                        Period2 = (item.Period2 + list2[i].Period2 + list3[i].Period2) * list_percent[i].Period2 + list_amount[i].Period2,
                        Period3 = (item.Period3 + list2[i].Period3 + list3[i].Period3) * list_percent[i].Period3 + list_amount[i].Period3,
                        Period4 = (item.Period4 + list2[i].Period4 + list3[i].Period4) * list_percent[i].Period4 + list_amount[i].Period4,
                        Period5 = (item.Period5 + list2[i].Period5 + list3[i].Period5) * list_percent[i].Period5 + list_amount[i].Period5,
                        Period6 = (item.Period6 + list2[i].Period6 + list3[i].Period6) * list_percent[i].Period6 + list_amount[i].Period6,
                        Period7 = (item.Period7 + list2[i].Period7 + list3[i].Period7) * list_percent[i].Period7 + list_amount[i].Period7,
                        Period8 = (item.Period8 + list2[i].Period8 + list3[i].Period8) * list_percent[i].Period8 + list_amount[i].Period8,
                        Period9 = (item.Period9 + list2[i].Period9 + list3[i].Period9) * list_percent[i].Period9 + list_amount[i].Period9,
                        Period10 = (item.Period10 + list2[i].Period10 + list3[i].Period10) * list_percent[i].Period10 + list_amount[i].Period10,
                        Period11 = (item.Period11 + list2[i].Period11 + list3[i].Period11) * list_percent[i].Period11 + list_amount[i].Period11,
                        Period12 = (item.Period12 + list2[i].Period12 + list3[i].Period12) * list_percent[i].Period12 + list_amount[i].Period12,
                        Period13 = (item.Period13 + list2[i].Period13 + list3[i].Period13) * list_percent[i].Period13 + list_amount[i].Period13,
                        Period14 = (item.Period14 + list2[i].Period14 + list3[i].Period14) * list_percent[i].Period14 + list_amount[i].Period14,
                        Period15 = (item.Period15 + list2[i].Period15 + list3[i].Period15) * list_percent[i].Period15 + list_amount[i].Period15,
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == type).ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }

        //4.1
        public string Calculate4_1(string version, string type)
        {
            List<FOL_Input_1_1> list1 = new List<FOL_Input_1_1>();
            List<FOL_Input_1_1> list2 = new List<FOL_Input_1_1>();
            List<FOL_Input_1_1> list3 = new List<FOL_Input_1_1>();
            list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            list3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_2_1> list_percent = dc.FOL_Input_2_1.Where(x => x.Type.Contains(type) && x.Version == version).ToList();
            List<FOL_Input_3_1> list_amount = dc.FOL_Input_3_1.Where(x => x.Type.Contains(type) && x.Version == version).ToList();
            string message = "";
            if (list1.Count == 0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list_percent.Count == 0)
            { message = message + "该月份的4.1 Material Loss (%)百分比数据为空;"; return message; }
            if (list_amount.Count == 0)
            { message = message + "该月份的4.1 Material Loss (%) Amount数据为空;"; return message; }
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int j = 0;
            try
            {
                for (int i = 0; i < list_percent.Count; i++)
                {
                    if (list_percent[i].GLOutputCode.Length > 15)
                    {
                        j = 0;
                        continue;
                    }
                    else
                    {
                        if (list_percent[i].GLOutputCode.Contains("AGI"))
                            j = 0;
                    }
                    FOL_Input_1_1 newItem = new FOL_Input_1_1();
                    newItem.GLBPCCode = list_percent[i].GLBPCCode;
                    newItem.GLBPCDescription = list_percent[i].GLBPCDescription;
                    newItem.GLOutputCode = list_percent[i].GLOutputCode;
                    newItem.CustomerBPCCode = list_percent[i].CustomerBPCCode;
                    newItem.DIM1 = list_percent[i].DIM1;
                    newItem.Order = list_percent[i].Order;
                    newItem.CustomerOutputCode = list_percent[i].CustomerOutputCode;
                    newItem.BU = list_percent[i].BU;
                    newItem.Segment = list_percent[i].Segment;
                    newItem.Period1 = (list1[j].Period1 + list2[j].Period1 + list3[j].Period1) * list_percent[i].Period1 + list_amount[i].Period1 ?? 0;
                    newItem.Period2 = (list1[j].Period2 + list2[j].Period2 + list3[j].Period2) * list_percent[i].Period2 + list_amount[i].Period2 ?? 0;
                    newItem.Period3 = (list1[j].Period3 + list2[j].Period3 + list3[j].Period3) * list_percent[i].Period3 + list_amount[i].Period3 ?? 0;
                    newItem.Period4 = (list1[j].Period4 + list2[j].Period4 + list3[j].Period4) * list_percent[i].Period4 + list_amount[i].Period4 ?? 0;
                    newItem.Period5 = (list1[j].Period5 + list2[j].Period5 + list3[j].Period5) * list_percent[i].Period5 + list_amount[i].Period5 ?? 0;
                    newItem.Period6 = (list1[j].Period6 + list2[j].Period6 + list3[j].Period6) * list_percent[i].Period6 + list_amount[i].Period6 ?? 0;
                    newItem.Period7 = (list1[j].Period7 + list2[j].Period7 + list3[j].Period7) * list_percent[i].Period7 + list_amount[i].Period7 ?? 0;
                    newItem.Period8 = (list1[j].Period8 + list2[j].Period8 + list3[j].Period8) * list_percent[i].Period8 + list_amount[i].Period8 ?? 0;
                    newItem.Period9 = (list1[j].Period9 + list2[j].Period9 + list3[j].Period9) * list_percent[i].Period9 + list_amount[i].Period9 ?? 0;
                    newItem.Period10 = (list1[j].Period10 + list2[j].Period10 + list3[j].Period10) * list_percent[i].Period10 + list_amount[i].Period10 ?? 0;
                    newItem.Period11 = (list1[j].Period11 + list2[j].Period11 + list3[j].Period11) * list_percent[i].Period11 + list_amount[i].Period11 ?? 0;
                    newItem.Period12 = (list1[j].Period12 + list2[j].Period12 + list3[j].Period12) * list_percent[i].Period12 + list_amount[i].Period12 ?? 0;
                    newItem.Period13 = (list1[j].Period13 + list2[j].Period13 + list3[j].Period13) * list_percent[i].Period13 + list_amount[i].Period13 ?? 0;
                    newItem.Period14 = (list1[j].Period14 + list2[j].Period14 + list3[j].Period14) * list_percent[i].Period14 + list_amount[i].Period14 ?? 0;
                    newItem.Period15 = (list1[j].Period15 + list2[j].Period15 + list3[j].Period15) * list_percent[i].Period15 + list_amount[i].Period15 ?? 0;
                    newItem.Period16 = (list1[j].Period16 + list2[j].Period16 + list3[j].Period16) * list_percent[i].Period16 + list_amount[i].Period16 ?? 0;
                    newItem.Type = type;
                    newItem.InsertDate = System.DateTime.Now;
                    newItem.InsertUser = "";
                    newItem.Version = version;
                    j++;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == type).ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";

            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }
        //5.1
        public string Calculate5_1(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_2_1> list3 = dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList();
            List<FOL_Input_3_1> list4 = dc.FOL_Input_3_1.Where(x => x.Type == type && x.Version == version).ToList();
            if (list1.Count == 0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的5.1 DL (%)百分比数据为空;"; return message; }
            if (list4.Count == 0)
            { message = message + "该月份的5.1 DL (%)Amount数据为空;"; return message; }

            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            int i = 0;
            int j = 0;
            try
            {
                foreach (var item in list4)
                {
                    if (j == 63)
                    {
                        j = 0;
                        continue;
                    }
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = item.GLBPCCode,
                        GLBPCDescription = item.GLBPCDescription,
                        GLOutputCode = item.GLOutputCode,
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = (list1[j].Period1 + list2[j].Period1) * list3[i].Period1 + list4[i].Period1,
                        Period2 = (list1[j].Period2 + list2[j].Period2) * list3[i].Period2 + list4[i].Period2,
                        Period3 = (list1[j].Period3 + list2[j].Period3) * list3[i].Period3 + list4[i].Period3,
                        Period4 = (list1[j].Period4 + list2[j].Period4) * list3[i].Period4 + list4[i].Period4,
                        Period5 = (list1[j].Period5 + list2[j].Period5) * list3[i].Period5 + list4[i].Period5,
                        Period6 = (list1[j].Period6 + list2[j].Period6) * list3[i].Period6 + list4[i].Period6,
                        Period7 = (list1[j].Period7 + list2[j].Period7) * list3[i].Period7 + list4[i].Period7,
                        Period8 = (list1[j].Period8 + list2[j].Period8) * list3[i].Period8 + list4[i].Period8,
                        Period9 = (list1[j].Period9 + list2[j].Period9) * list3[i].Period9 + list4[i].Period9,
                        Period10 = (list1[j].Period10 + list2[j].Period10) * list3[i].Period10 + list4[i].Period10,
                        Period11 = (list1[j].Period11 + list2[j].Period11) * list3[i].Period11 + list4[i].Period11,
                        Period12 = (list1[j].Period12 + list2[j].Period12) * list3[i].Period12 + list4[i].Period12,
                        Period13 = (list1[j].Period13 + list2[j].Period13) * list3[i].Period13 + list4[i].Period13,
                        Period14 = (list1[j].Period14 + list2[j].Period14) * list3[i].Period14 + list4[i].Period14,
                        Period15 = (list1[j].Period15 + list2[j].Period15) * list3[i].Period15 + list4[i].Period15,
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                    i++;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "5.1 DL (%)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }

            return message;
        }

        //6.0 7.0
        public string Calculate6_0(string version, string type, string department)
        {
            string message = "";
            List<FOL_Input_6_0_TotalMOH> list = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department == department && x.Version == version && x.Type == type).ToList();
            List<FOL_Input_6_0_TotalMOH> list_Sum = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department == department && x.Version == version && x.Type == type && x.IsSubtotal == "11").ToList();
            try
            {
                for (int i = 0; i < list.Count; i++)
                {
                    if (list[i].IsSubtotal != "" && list[i].IsSubtotal != null && list[i].IsSubtotal != "11")
                    {
                        list[i].Period1 = 0;
                        list[i].Period2 = 0;
                        list[i].Period3 = 0;
                        list[i].Period4 = 0;
                        list[i].Period5 = 0;
                        list[i].Period6 = 0;
                        list[i].Period7 = 0;
                        list[i].Period8 = 0;
                        list[i].Period9 = 0;
                        list[i].Period10 = 0;
                        list[i].Period11 = 0;
                        list[i].Period12 = 0;
                        list[i].Period13 = 0;
                        list[i].Period14 = 0;
                        list[i].Period15 = 0;
                        list[i].Period16 = 0;
                        int start = Convert.ToInt32(list[i].IsSubtotal.Split(':')[0]);
                        int end = Convert.ToInt32(list[i].IsSubtotal.Split(':')[1]);
                        for (int n = start - 1; n < end; n++)
                        {
                            list[i].Period1 += list[n].Period1;
                            list[i].Period2 += list[n].Period2;
                            list[i].Period3 += list[n].Period3;
                            list[i].Period4 += list[n].Period4;
                            list[i].Period5 += list[n].Period5;
                            list[i].Period6 += list[n].Period6;
                            list[i].Period7 += list[n].Period7;
                            list[i].Period8 += list[n].Period8;
                            list[i].Period9 += list[n].Period9;
                            list[i].Period10 += list[n].Period10;
                            list[i].Period11 += list[n].Period11;
                            list[i].Period12 += list[n].Period12;
                            list[i].Period13 += list[n].Period13;
                            list[i].Period14 += list[n].Period14;
                            list[i].Period15 += list[n].Period15;
                            list[i].Period16 += list[n].Period16;
                        }
                    }
                }
                for (int i = 0; i < list_Sum.Count; i++)
                {
                    list_Sum[i].Period1 = 0;
                    list_Sum[i].Period2 = 0;
                    list_Sum[i].Period3 = 0;
                    list_Sum[i].Period4 = 0;
                    list_Sum[i].Period5 = 0;
                    list_Sum[i].Period6 = 0;
                    list_Sum[i].Period7 = 0;
                    list_Sum[i].Period8 = 0;
                    list_Sum[i].Period9 = 0;
                    list_Sum[i].Period10 = 0;
                    list_Sum[i].Period11 = 0;
                    list_Sum[i].Period12 = 0;
                    list_Sum[i].Period13 = 0;
                    list_Sum[i].Period14 = 0;
                    list_Sum[i].Period15 = 0;
                    list_Sum[i].Period16 = 0;
                    for (int j = 0; j < list.Count; j++)
                    {
                        if (list[j].BPCOutputCode == list_Sum[i].PARENTH1)
                        {
                            list_Sum[i].Period1 += list[j].Period1;
                            list_Sum[i].Period2 += list[j].Period2;
                            list_Sum[i].Period3 += list[j].Period3;
                            list_Sum[i].Period4 += list[j].Period4;
                            list_Sum[i].Period5 += list[j].Period5;
                            list_Sum[i].Period6 += list[j].Period6;
                            list_Sum[i].Period7 += list[j].Period7;
                            list_Sum[i].Period8 += list[j].Period8;
                            list_Sum[i].Period9 += list[j].Period9;
                            list_Sum[i].Period10 += list[j].Period10;
                            list_Sum[i].Period11 += list[j].Period11;
                            list_Sum[i].Period12 += list[j].Period12;
                            list_Sum[i].Period13 += list[j].Period13;
                            list_Sum[i].Period14 += list[j].Period14;
                            list_Sum[i].Period15 += list[j].Period15;
                            list_Sum[i].Period16 += list[j].Period16;
                        }
                    }
                }
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {

                message = ex.Message;
            }
            return message;
        }

        //6.1
        public string Calculate6_1(string version, string type, string cc)
        {
            string message = "";
            List<FOL_Input_2_1> list_percent = new List<FOL_Input_2_1>();
            List<FOL_Input_1_1> list_forecast = new List<FOL_Input_1_1>();

            List<FOL_Input_3_1> list_amount = dc.FOL_Input_3_1.Where(x => x.Version == version && x.Type.Contains("6.0") && x.Department == cc).ToList();
            if (list_amount == null) { message = "该月份的6.1 IDL (%&#) Amount数据为空;"; return message; }

            //找到6.0这张表中IDL的Total值
            FOL_Input_6_0_TotalMOH totalIDL = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Version == version && x.Type == type && x.Department == cc && x.BPCOutputCode == "4606-IDL").FirstOrDefault();
            if (totalIDL == null) { message = "该月份的6.0_TotalMOH数据为空;"; return message; }

            try
            {
                FOL_Input_3_1 total_amount = new FOL_Input_3_1();
                for (int i = 0; i < list_amount.Count; i++)
                {
                    total_amount.Period1 = (total_amount.Period1 == null ? 0 : total_amount.Period1) + list_amount[i].Period1;
                    total_amount.Period2 = (total_amount.Period2 == null ? 0 : total_amount.Period2) + list_amount[i].Period2;
                    total_amount.Period3 = (total_amount.Period3 == null ? 0 : total_amount.Period3) + list_amount[i].Period3;
                    total_amount.Period4 = (total_amount.Period4 == null ? 0 : total_amount.Period4) + list_amount[i].Period4;
                    total_amount.Period5 = (total_amount.Period5 == null ? 0 : total_amount.Period5) + list_amount[i].Period5;
                    total_amount.Period6 = (total_amount.Period6 == null ? 0 : total_amount.Period6) + list_amount[i].Period6;
                    total_amount.Period7 = (total_amount.Period7 == null ? 0 : total_amount.Period7) + list_amount[i].Period7;
                    total_amount.Period8 = (total_amount.Period8 == null ? 0 : total_amount.Period8) + list_amount[i].Period8;
                    total_amount.Period9 = (total_amount.Period9 == null ? 0 : total_amount.Period9) + list_amount[i].Period9;
                    total_amount.Period10 = (total_amount.Period10 == null ? 0 : total_amount.Period10) + list_amount[i].Period10;
                    total_amount.Period11 = (total_amount.Period11 == null ? 0 : total_amount.Period11) + list_amount[i].Period11;
                    total_amount.Period12 = (total_amount.Period12 == null ? 0 : total_amount.Period12) + list_amount[i].Period12;
                    total_amount.Period13 = (total_amount.Period13 == null ? 0 : total_amount.Period13) + list_amount[i].Period13;
                    total_amount.Period14 = (total_amount.Period14 == null ? 0 : total_amount.Period14) + list_amount[i].Period14;
                    total_amount.Period15 = (total_amount.Period15 == null ? 0 : total_amount.Period15) + list_amount[i].Period15;
                    total_amount.Period16 = (total_amount.Period16 == null ? 0 : total_amount.Period16) + list_amount[i].Period16;
                }

                for (int i = 0; i < list_amount.Count; i++)
                {
                    FOL_Input_2_1 newItemPercent = new FOL_Input_2_1();
                    FOL_Input_1_1 newItemForecast = new FOL_Input_1_1();

                    //计算Percent
                    #region
                    newItemPercent.GLBPCCode = list_amount[i].GLBPCCode;
                    newItemPercent.GLBPCDescription = list_amount[i].GLBPCDescription;
                    newItemPercent.GLOutputCode = list_amount[i].GLOutputCode;
                    newItemPercent.CustomerBPCCode = list_amount[i].CustomerBPCCode;
                    newItemPercent.DIM1 = list_amount[i].DIM1;
                    newItemPercent.Order = list_amount[i].Order;
                    newItemPercent.CustomerOutputCode = list_amount[i].CustomerOutputCode;
                    newItemPercent.BU = list_amount[i].BU;
                    newItemPercent.Segment = list_amount[i].Segment;
                    newItemPercent.Period1 = (total_amount.Period1 == 0) ? 0 : (list_amount[i].Period1 / total_amount.Period1);
                    newItemPercent.Period2 = (total_amount.Period2 == 0) ? 0 : (list_amount[i].Period2 / total_amount.Period2);
                    newItemPercent.Period3 = (total_amount.Period3 == 0) ? 0 : (list_amount[i].Period3 / total_amount.Period3);
                    newItemPercent.Period4 = (total_amount.Period4 == 0) ? 0 : (list_amount[i].Period4 / total_amount.Period4);
                    newItemPercent.Period5 = (total_amount.Period5 == 0) ? 0 : (list_amount[i].Period5 / total_amount.Period5);
                    newItemPercent.Period6 = (total_amount.Period6 == 0) ? 0 : (list_amount[i].Period6 / total_amount.Period6);
                    newItemPercent.Period7 = (total_amount.Period7 == 0) ? 0 : (list_amount[i].Period7 / total_amount.Period7);
                    newItemPercent.Period8 = (total_amount.Period8 == 0) ? 0 : (list_amount[i].Period8 / total_amount.Period8);
                    newItemPercent.Period9 = (total_amount.Period9 == 0) ? 0 : (list_amount[i].Period9 / total_amount.Period9);
                    newItemPercent.Period10 = (total_amount.Period10 == 0) ? 0 : (list_amount[i].Period10 / total_amount.Period10);
                    newItemPercent.Period11 = (total_amount.Period11 == 0) ? 0 : (list_amount[i].Period11 / total_amount.Period11);
                    newItemPercent.Period12 = (total_amount.Period12 == 0) ? 0 : (list_amount[i].Period12 / total_amount.Period12);
                    newItemPercent.Period13 = (total_amount.Period13 == 0) ? 0 : (list_amount[i].Period13 / total_amount.Period13);
                    newItemPercent.Period14 = (total_amount.Period14 == 0) ? 0 : (list_amount[i].Period14 / total_amount.Period14);
                    newItemPercent.Period15 = (total_amount.Period15 == 0) ? 0 : (list_amount[i].Period15 / total_amount.Period15);
                    newItemPercent.Period16 = (total_amount.Period16 == 0) ? 0 : (list_amount[i].Period16 / total_amount.Period16);
                    newItemPercent.Type = type;
                    newItemPercent.InsertDate = System.DateTime.Now;
                    newItemPercent.InsertUser = "";
                    newItemPercent.Version = version;
                    newItemPercent.Department = list_amount[i].Department;
                    list_percent.Add(newItemPercent);
                    #endregion

                    //计算Dollar Amount
                    #region
                    newItemForecast.GLBPCCode = list_amount[i].GLBPCCode;
                    newItemForecast.GLBPCDescription = list_amount[i].GLBPCDescription;
                    newItemForecast.GLOutputCode = list_amount[i].GLOutputCode;
                    newItemForecast.CustomerBPCCode = list_amount[i].CustomerBPCCode;
                    newItemForecast.DIM1 = list_amount[i].DIM1;
                    newItemForecast.Order = list_amount[i].Order;
                    newItemForecast.CustomerOutputCode = list_amount[i].CustomerOutputCode;
                    newItemForecast.BU = list_amount[i].BU;
                    newItemForecast.Segment = list_amount[i].Segment;
                    newItemForecast.Period1 = newItemPercent.Period1 * totalIDL.Period1;
                    newItemForecast.Period2 = newItemPercent.Period2 * totalIDL.Period2;
                    newItemForecast.Period3 = newItemPercent.Period3 * totalIDL.Period3;
                    newItemForecast.Period4 = newItemPercent.Period4 * totalIDL.Period4;
                    newItemForecast.Period5 = newItemPercent.Period5 * totalIDL.Period5;
                    newItemForecast.Period6 = newItemPercent.Period6 * totalIDL.Period6;
                    newItemForecast.Period7 = newItemPercent.Period7 * totalIDL.Period7;
                    newItemForecast.Period8 = newItemPercent.Period8 * totalIDL.Period8;
                    newItemForecast.Period9 = newItemPercent.Period9 * totalIDL.Period9;
                    newItemForecast.Period10 = newItemPercent.Period10 * totalIDL.Period10;
                    newItemForecast.Period11 = newItemPercent.Period11 * totalIDL.Period11;
                    newItemForecast.Period12 = newItemPercent.Period12 * totalIDL.Period12;
                    newItemForecast.Period13 = newItemPercent.Period13 * totalIDL.Period13;
                    newItemForecast.Period14 = newItemPercent.Period14 * totalIDL.Period14;
                    newItemForecast.Period15 = newItemPercent.Period15 * totalIDL.Period15;
                    newItemForecast.Period16 = newItemPercent.Period16 * totalIDL.Period16;
                    newItemForecast.Type = type;
                    newItemForecast.InsertDate = System.DateTime.Now;
                    newItemForecast.InsertUser = "";
                    newItemForecast.Version = version;
                    newItemForecast.Department = list_amount[i].Department;
                    list_forecast.Add(newItemForecast);
                    #endregion
                }

                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Version == version && x.Type == type && x.Department == cc));
                dc.FOL_Input_2_1.InsertAllOnSubmit(list_percent);
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == type && x.Department == cc));
                dc.FOL_Input_1_1.InsertAllOnSubmit(list_forecast);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        //6.2, 6.3, 6.4, 6.5, 6.6
        public string Calculate6_2(string version, string type)
        {
            string message = "";
            List<FOL_Input_2_1> list = new List<FOL_Input_2_1>();
            List<FOL_Input_1_1> list_Forecast = new List<FOL_Input_1_1>();

            List<FOL_Input_3_1> list_amount = dc.FOL_Input_3_1.Where(x => x.Version == version && x.Type == type).ToList();
            if (list_amount == null) { message = "该月份的" + type + "数据为空;"; return message; }
            string parenth1 = "";
            switch (type)
            {
                case "6.2 Depn(%)":
                    parenth1 = "4607";
                    break;
                case "6.3 FUS.(%)":
                    parenth1 = "4608";
                    break;
                case "6.4 MOHO(%)":
                    parenth1 = "4610";
                    break;
                case "6.5 OEUP(%)":
                    parenth1 = "4611";
                    break;
                case "6.6 ELEC(%)":
                    parenth1 = "4612";
                    break;
                default:
                    parenth1 = "";
                    break;
            }
            List<FOL_Input_6_0_TotalMOH> list_6_0 = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Version == version && x.Type.Contains("6.0") && x.IsSubtotal == "11" && x.PARENTH1 == parenth1).ToList();
            if (list_6_0 == null) { message = "该月份的6.0 TotalMoht数据为空;"; return message; }
            FOL_Input_6_0_TotalMOH total_6_0 = new FOL_Input_6_0_TotalMOH();
            total_6_0.Period1 = 0;
            total_6_0.Period2 = 0;
            total_6_0.Period3 = 0;
            total_6_0.Period4 = 0;
            total_6_0.Period5 = 0;
            total_6_0.Period6 = 0;
            total_6_0.Period7 = 0;
            total_6_0.Period8 = 0;
            total_6_0.Period9 = 0;
            total_6_0.Period10 = 0;
            total_6_0.Period11 = 0;
            total_6_0.Period12 = 0;
            total_6_0.Period13 = 0;
            total_6_0.Period14 = 0;
            total_6_0.Period15 = 0;
            total_6_0.Period16 = 0;
            for (int i = 0; i < list_6_0.Count; i++)
            {
                total_6_0.Period1 = +list_6_0[i].Period1;
                total_6_0.Period2 = +list_6_0[i].Period2;
                total_6_0.Period3 = +list_6_0[i].Period3;
                total_6_0.Period4 = +list_6_0[i].Period4;
                total_6_0.Period5 = +list_6_0[i].Period5;
                total_6_0.Period6 = +list_6_0[i].Period6;
                total_6_0.Period7 = +list_6_0[i].Period7;
                total_6_0.Period8 = +list_6_0[i].Period8;
                total_6_0.Period9 = +list_6_0[i].Period9;
                total_6_0.Period10 = +list_6_0[i].Period10;
                total_6_0.Period11 = +list_6_0[i].Period11;
                total_6_0.Period12 = +list_6_0[i].Period12;
                total_6_0.Period13 = +list_6_0[i].Period13;
                total_6_0.Period14 = +list_6_0[i].Period14;
                total_6_0.Period15 = +list_6_0[i].Period15;
                total_6_0.Period16 = +list_6_0[i].Period16;
            }
            try
            {
                FOL_Input_3_1 total_amount = new FOL_Input_3_1();
                for (int i = 0; i < list_amount.Count; i++)
                {
                    total_amount.Period1 = (total_amount.Period1 == null ? 0 : total_amount.Period1) + list_amount[i].Period1;
                    total_amount.Period2 = (total_amount.Period2 == null ? 0 : total_amount.Period2) + list_amount[i].Period2;
                    total_amount.Period3 = (total_amount.Period3 == null ? 0 : total_amount.Period3) + list_amount[i].Period3;
                    total_amount.Period4 = (total_amount.Period4 == null ? 0 : total_amount.Period4) + list_amount[i].Period4;
                    total_amount.Period5 = (total_amount.Period5 == null ? 0 : total_amount.Period5) + list_amount[i].Period5;
                    total_amount.Period6 = (total_amount.Period6 == null ? 0 : total_amount.Period6) + list_amount[i].Period6;
                    total_amount.Period7 = (total_amount.Period7 == null ? 0 : total_amount.Period7) + list_amount[i].Period7;
                    total_amount.Period8 = (total_amount.Period8 == null ? 0 : total_amount.Period8) + list_amount[i].Period8;
                    total_amount.Period9 = (total_amount.Period9 == null ? 0 : total_amount.Period9) + list_amount[i].Period9;
                    total_amount.Period10 = (total_amount.Period10 == null ? 0 : total_amount.Period10) + list_amount[i].Period10;
                    total_amount.Period11 = (total_amount.Period11 == null ? 0 : total_amount.Period11) + list_amount[i].Period11;
                    total_amount.Period12 = (total_amount.Period12 == null ? 0 : total_amount.Period12) + list_amount[i].Period12;
                    total_amount.Period13 = (total_amount.Period13 == null ? 0 : total_amount.Period13) + list_amount[i].Period13;
                    total_amount.Period14 = (total_amount.Period14 == null ? 0 : total_amount.Period14) + list_amount[i].Period14;
                    total_amount.Period15 = (total_amount.Period15 == null ? 0 : total_amount.Period15) + list_amount[i].Period15;
                    total_amount.Period16 = (total_amount.Period16 == null ? 0 : total_amount.Period16) + list_amount[i].Period16;
                }
                #region
                for (int i = 0; i < list_amount.Count; i++)
                {
                    FOL_Input_2_1 newItem_Percent = new FOL_Input_2_1();
                    FOL_Input_1_1 newItem_Forecast = new FOL_Input_1_1();

                    //计算percent
                    newItem_Percent.GLBPCCode = list_amount[i].GLBPCCode;
                    newItem_Percent.GLBPCDescription = list_amount[i].GLBPCDescription;
                    newItem_Percent.GLOutputCode = list_amount[i].GLOutputCode;
                    newItem_Percent.CustomerBPCCode = list_amount[i].CustomerBPCCode;
                    newItem_Percent.DIM1 = list_amount[i].DIM1;
                    newItem_Percent.Order = list_amount[i].Order;
                    newItem_Percent.CustomerOutputCode = list_amount[i].CustomerOutputCode;
                    newItem_Percent.BU = list_amount[i].BU;
                    newItem_Percent.Segment = list_amount[i].Segment;
                    newItem_Percent.Period1 = (total_amount.Period1 == 0) ? 0 : (list_amount[i].Period1 / total_amount.Period1);
                    newItem_Percent.Period2 = (total_amount.Period2 == 0) ? 0 : (list_amount[i].Period2 / total_amount.Period2);
                    newItem_Percent.Period3 = (total_amount.Period3 == 0) ? 0 : (list_amount[i].Period3 / total_amount.Period3);
                    newItem_Percent.Period4 = (total_amount.Period4 == 0) ? 0 : (list_amount[i].Period4 / total_amount.Period4);
                    newItem_Percent.Period5 = (total_amount.Period5 == 0) ? 0 : (list_amount[i].Period5 / total_amount.Period5);
                    newItem_Percent.Period6 = (total_amount.Period6 == 0) ? 0 : (list_amount[i].Period6 / total_amount.Period6);
                    newItem_Percent.Period7 = (total_amount.Period7 == 0) ? 0 : (list_amount[i].Period7 / total_amount.Period7);
                    newItem_Percent.Period8 = (total_amount.Period8 == 0) ? 0 : (list_amount[i].Period8 / total_amount.Period8);
                    newItem_Percent.Period9 = (total_amount.Period9 == 0) ? 0 : (list_amount[i].Period9 / total_amount.Period9);
                    newItem_Percent.Period10 = (total_amount.Period10 == 0) ? 0 : (list_amount[i].Period10 / total_amount.Period10);
                    newItem_Percent.Period11 = (total_amount.Period11 == 0) ? 0 : (list_amount[i].Period11 / total_amount.Period11);
                    newItem_Percent.Period12 = (total_amount.Period12 == 0) ? 0 : (list_amount[i].Period12 / total_amount.Period12);
                    newItem_Percent.Period13 = (total_amount.Period13 == 0) ? 0 : (list_amount[i].Period13 / total_amount.Period13);
                    newItem_Percent.Period14 = (total_amount.Period14 == 0) ? 0 : (list_amount[i].Period14 / total_amount.Period14);
                    newItem_Percent.Period15 = (total_amount.Period15 == 0) ? 0 : (list_amount[i].Period15 / total_amount.Period15);
                    newItem_Percent.Period16 = (total_amount.Period16 == 0) ? 0 : (list_amount[i].Period16 / total_amount.Period16);
                    newItem_Percent.Type = type;
                    newItem_Percent.InsertDate = System.DateTime.Now;
                    newItem_Percent.InsertUser = "";
                    newItem_Percent.Version = version;
                    newItem_Percent.Department = "";
                    list.Add(newItem_Percent);

                    //计算forecast
                    newItem_Forecast.GLBPCCode = list_amount[i].GLBPCCode;
                    newItem_Forecast.GLBPCDescription = list_amount[i].GLBPCDescription;
                    newItem_Forecast.GLOutputCode = list_amount[i].GLOutputCode;
                    newItem_Forecast.CustomerBPCCode = list_amount[i].CustomerBPCCode;
                    newItem_Forecast.DIM1 = list_amount[i].DIM1;
                    newItem_Forecast.Order = list_amount[i].Order;
                    newItem_Forecast.CustomerOutputCode = list_amount[i].CustomerOutputCode;
                    newItem_Forecast.BU = list_amount[i].BU;
                    newItem_Forecast.Segment = list_amount[i].Segment;
                    newItem_Forecast.Period1 = total_6_0.Period1 * newItem_Percent.Period1;
                    newItem_Forecast.Period2 = total_6_0.Period2 * newItem_Percent.Period2;
                    newItem_Forecast.Period3 = total_6_0.Period3 * newItem_Percent.Period3;
                    newItem_Forecast.Period4 = total_6_0.Period4 * newItem_Percent.Period4;
                    newItem_Forecast.Period5 = total_6_0.Period5 * newItem_Percent.Period5;
                    newItem_Forecast.Period6 = total_6_0.Period6 * newItem_Percent.Period6;
                    newItem_Forecast.Period7 = total_6_0.Period7 * newItem_Percent.Period7;
                    newItem_Forecast.Period8 = total_6_0.Period8 * newItem_Percent.Period8;
                    newItem_Forecast.Period9 = total_6_0.Period9 * newItem_Percent.Period9;
                    newItem_Forecast.Period10 = total_6_0.Period10 * newItem_Percent.Period10;
                    newItem_Forecast.Period11 = total_6_0.Period11 * newItem_Percent.Period11;
                    newItem_Forecast.Period12 = total_6_0.Period12 * newItem_Percent.Period12;
                    newItem_Forecast.Period13 = total_6_0.Period13 * newItem_Percent.Period13;
                    newItem_Forecast.Period14 = total_6_0.Period14 * newItem_Percent.Period14;
                    newItem_Forecast.Period15 = total_6_0.Period15 * newItem_Percent.Period15;
                    newItem_Forecast.Period16 = total_6_0.Period16 * newItem_Percent.Period16;
                    newItem_Forecast.Version = version;
                    newItem_Forecast.Type = type;
                    newItem_Forecast.InsertDate = System.DateTime.Now;
                    newItem_Forecast.InsertUser = "";
                    newItem_Forecast.Department = "";
                    list_Forecast.Add(newItem_Forecast);
                }
                #endregion
                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList());
                dc.FOL_Input_2_1.InsertAllOnSubmit(list);
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Type == type && x.Version == version).ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list_Forecast);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        //计算7.1，7.2
        public string Calculate7_1(string version, string type, string cc)
        {
            string message = "";
            FOL_Input_6_0_TotalMOH item_total = new FOL_Input_6_0_TotalMOH();

            //get total
            #region
            if (type.Contains("7.1"))
            {
                List<FOL_Input_6_0_TotalMOH> list_6_0 = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Version == version && x.Type.Contains("7.0") && x.IsSubtotal == "11"
                  && (x.Department.ToUpper() == "CC61" || x.Department.ToUpper() == "CC62" || x.Department.ToUpper() == "CC63") && x.GLBPCCode == cc).ToList();
                FOL_Input_6_0_TotalMOH item = new FOL_Input_6_0_TotalMOH();
                for (int n = 0; n < list_6_0.Count(); n++)
                {
                    if (list_6_0[n].GLBPCCode == cc)
                    {
                        item.Period1 = (item.Period1 == null ? 0 : item.Period1) + list_6_0[n].Period1;
                        item.Period2 = (item.Period2 == null ? 0 : item.Period2) + list_6_0[n].Period2;
                        item.Period3 = (item.Period3 == null ? 0 : item.Period3) + list_6_0[n].Period3;
                        item.Period4 = (item.Period4 == null ? 0 : item.Period4) + list_6_0[n].Period4;
                        item.Period5 = (item.Period5 == null ? 0 : item.Period5) + list_6_0[n].Period5;
                        item.Period6 = (item.Period6 == null ? 0 : item.Period6) + list_6_0[n].Period6;
                        item.Period7 = (item.Period7 == null ? 0 : item.Period7) + list_6_0[n].Period7;
                        item.Period8 = (item.Period8 == null ? 0 : item.Period8) + list_6_0[n].Period8;
                        item.Period9 = (item.Period9 == null ? 0 : item.Period9) + list_6_0[n].Period9;
                        item.Period10 = (item.Period10 == null ? 0 : item.Period10) + list_6_0[n].Period10;
                        item.Period11 = (item.Period11 == null ? 0 : item.Period11) + list_6_0[n].Period11;
                        item.Period12 = (item.Period12 == null ? 0 : item.Period12) + list_6_0[n].Period12;
                        item.Period13 = (item.Period13 == null ? 0 : item.Period13) + list_6_0[n].Period13;
                        item.Period14 = (item.Period14 == null ? 0 : item.Period14) + list_6_0[n].Period14;
                        item.Period15 = (item.Period15 == null ? 0 : item.Period15) + list_6_0[n].Period15;
                        item.Period16 = (item.Period16 == null ? 0 : item.Period16) + list_6_0[n].Period16;
                    }
                }
            }
            else
            {
                item_total = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Version == version && x.Type.Contains("7.0") && x.IsSubtotal == "11"
                   && x.Department.ToUpper() == "CCSM" && x.GLBPCCode == cc).FirstOrDefault();
            }
            #endregion
            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            message = Calculate7_1_Percent(version, type);
            if (message != "Success") { return message; }
            List<FOL_Input_2_1> list_percent = dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList();
            string GLDescription = dc.FOL_ExpenseMappingManagement_.Where(x => x.GLBPCCode == cc).FirstOrDefault().GLBPCDescription;
            try
            {
                for (int i = 0; i < list_percent.Count; i++)
                {
                    FOL_Input_1_1 newItem = new FOL_Input_1_1();
                    newItem.GLBPCCode = cc;
                    newItem.GLBPCDescription = GLDescription;
                    newItem.GLOutputCode = list_percent[i].GLOutputCode;
                    newItem.CustomerBPCCode = list_percent[i].CustomerBPCCode;
                    newItem.DIM1 = list_percent[i].DIM1;
                    newItem.Order = list_percent[i].Order;
                    newItem.CustomerOutputCode = list_percent[i].CustomerOutputCode;
                    newItem.BU = list_percent[i].BU;
                    newItem.Segment = list_percent[i].Segment;
                    newItem.Period1 = item_total.Period1 * (list_percent[i].Period1 ?? 0);
                    newItem.Period2 = item_total.Period2 * (list_percent[i].Period2 ?? 0);
                    newItem.Period3 = item_total.Period3 * (list_percent[i].Period3 ?? 0);
                    newItem.Period4 = item_total.Period4 * (list_percent[i].Period4 ?? 0);
                    newItem.Period5 = item_total.Period5 * (list_percent[i].Period5 ?? 0);
                    newItem.Period6 = item_total.Period6 * (list_percent[i].Period6 ?? 0);
                    newItem.Period7 = item_total.Period7 * (list_percent[i].Period7 ?? 0);
                    newItem.Period8 = item_total.Period8 * (list_percent[i].Period8 ?? 0);
                    newItem.Period9 = item_total.Period9 * (list_percent[i].Period9 ?? 0);
                    newItem.Period10 = item_total.Period10 * (list_percent[i].Period10 ?? 0);
                    newItem.Period11 = item_total.Period11 * (list_percent[i].Period11 ?? 0);
                    newItem.Period12 = item_total.Period12 * (list_percent[i].Period12 ?? 0);
                    newItem.Period13 = item_total.Period13 * (list_percent[i].Period13 ?? 0);
                    newItem.Period14 = item_total.Period14 * (list_percent[i].Period14 ?? 0);
                    newItem.Period15 = item_total.Period15 * (list_percent[i].Period15 ?? 0);
                    newItem.Period16 = item_total.Period16 * (list_percent[i].Period16 ?? 0);
                    newItem.Type = type;
                    newItem.Version = version;
                    newItem.InsertDate = DateTime.Now;
                    newItem.InsertUser = "";
                    newItem.Department = cc;
                    list.Add(newItem);
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Type == type && x.Version == version && x.Department == cc).ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        //计算7.1, 7.2的percent
        public string Calculate7_1_Percent(string version, string type)
        {
            string message = "";
            //Caculate percent
            List<FOL_Input_1_1> projectList1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            if (projectList1.Count() == 0) { message = "1.1 is empty!"; return message; }
            List<FOL_Input_1_1> projectList2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            if (projectList2.Count() == 0) { message = "1.2 is empty!"; return message; }
            List<FOL_Input_1_1> projectList3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            if (projectList3.Count() == 0) { message = "1.3 is empty!"; return message; }
            List<FOL_Input_2_1> percentList2_1 = dc.FOL_Input_2_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            if (percentList2_1.Count() == 0) { message = "2.1 percent is empty!"; return message; }
            List<FOL_Input_1_1> projectList2_1 = dc.FOL_Input_1_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            if (projectList2_1.Count() == 0) { message = "2.1 forecast is empty!"; return message; }
            List<FOL_Input_1_1> projectList1_0 = dc.FOL_Input_1_1.Where(x => x.Type == "1.0 Total Sales($)" && x.Version == version).ToList();
            if (projectList1_0.Count() == 0) { message = "1.0 is empty!"; return message; }
            FOL_Input_1_1 site1 = new FOL_Input_1_1();
            FOL_Input_1_1 site2 = new FOL_Input_1_1();
            FOL_Input_1_1 site3 = new FOL_Input_1_1();
            FOL_Input_1_1 site1_0 = new FOL_Input_1_1();
            FOL_Input_1_1 site2_1 = new FOL_Input_1_1();
            FOL_Input_2_1 sitePercent2_1 = new FOL_Input_2_1();


            List<FOL_Input_1_1> List_ProjectSumOfThree = new List<FOL_Input_1_1>();
            FOL_Input_1_1 List_SiteSumOfThree = new FOL_Input_1_1();
            List<FOL_Input_1_1> List_ProjectSumOfTwo = new List<FOL_Input_1_1>();
            FOL_Input_1_1 List_SiteSumOfTwo = new FOL_Input_1_1();

            List<FOL_Input_2_1> listResult = new List<FOL_Input_2_1>();
            //********计算 site值
            #region
            for (int i = 0; i < projectList1.Count; i++)
            {
                site1.Period1 = +(site1.Period1 ?? 0) + projectList1[i].Period1;
                site1.Period2 = +(site1.Period2 ?? 0) + projectList1[i].Period2;
                site1.Period3 = +(site1.Period3 ?? 0) + projectList1[i].Period3;
                site1.Period4 = +(site1.Period4 ?? 0) + projectList1[i].Period4;
                site1.Period5 = +(site1.Period5 ?? 0) + projectList1[i].Period5;
                site1.Period6 = +(site1.Period6 ?? 0) + projectList1[i].Period6;
                site1.Period7 = +(site1.Period7 ?? 0) + projectList1[i].Period7;
                site1.Period8 = +(site1.Period8 ?? 0) + projectList1[i].Period8;
                site1.Period9 = +(site1.Period9 ?? 0) + projectList1[i].Period9;
                site1.Period10 = +(site1.Period10 ?? 0) + projectList1[i].Period10;
                site1.Period11 = +(site1.Period11 ?? 0) + projectList1[i].Period11;
                site1.Period12 = +(site1.Period12 ?? 0) + projectList1[i].Period12;
                site1.Period13 = +(site1.Period13 ?? 0) + projectList1[i].Period13;
                site1.Period14 = +(site1.Period14 ?? 0) + projectList1[i].Period14;
                site1.Period15 = +(site1.Period15 ?? 0) + projectList1[i].Period15;
                site1.Period16 = +(site1.Period16 ?? 0) + projectList1[i].Period16;

                site2.Period1 = +(site2.Period1 ?? 0) + projectList2[i].Period1;
                site2.Period2 = +(site2.Period2 ?? 0) + projectList2[i].Period2;
                site2.Period3 = +(site2.Period3 ?? 0) + projectList2[i].Period3;
                site2.Period4 = +(site2.Period4 ?? 0) + projectList2[i].Period4;
                site2.Period5 = +(site2.Period5 ?? 0) + projectList2[i].Period5;
                site2.Period6 = +(site2.Period6 ?? 0) + projectList2[i].Period6;
                site2.Period7 = +(site2.Period7 ?? 0) + projectList2[i].Period7;
                site2.Period8 = +(site2.Period8 ?? 0) + projectList2[i].Period8;
                site2.Period9 = +(site2.Period9 ?? 0) + projectList2[i].Period9;
                site2.Period10 = +(site2.Period10 ?? 0) + projectList2[i].Period10;
                site2.Period11 = +(site2.Period11 ?? 0) + projectList2[i].Period11;
                site2.Period12 = +(site2.Period12 ?? 0) + projectList2[i].Period12;
                site2.Period13 = +(site2.Period13 ?? 0) + projectList2[i].Period13;
                site2.Period14 = +(site2.Period14 ?? 0) + projectList2[i].Period14;
                site2.Period15 = +(site2.Period15 ?? 0) + projectList2[i].Period15;
                site2.Period16 = +(site2.Period16 ?? 0) + projectList2[i].Period16;

                site3.Period1 = +(site3.Period1 ?? 0) + projectList3[i].Period1;
                site3.Period2 = +(site3.Period2 ?? 0) + projectList3[i].Period2;
                site3.Period3 = +(site3.Period3 ?? 0) + projectList3[i].Period3;
                site3.Period4 = +(site3.Period4 ?? 0) + projectList3[i].Period4;
                site3.Period5 = +(site3.Period5 ?? 0) + projectList3[i].Period5;
                site3.Period6 = +(site3.Period6 ?? 0) + projectList3[i].Period6;
                site3.Period7 = +(site3.Period7 ?? 0) + projectList3[i].Period7;
                site3.Period8 = +(site3.Period8 ?? 0) + projectList3[i].Period8;
                site3.Period9 = +(site3.Period9 ?? 0) + projectList3[i].Period9;
                site3.Period10 = +(site3.Period10 ?? 0) + projectList3[i].Period10;
                site3.Period11 = +(site3.Period11 ?? 0) + projectList3[i].Period11;
                site3.Period12 = +(site3.Period12 ?? 0) + projectList3[i].Period12;
                site3.Period13 = +(site3.Period13 ?? 0) + projectList3[i].Period13;
                site3.Period14 = +(site3.Period14 ?? 0) + projectList3[i].Period14;
                site3.Period15 = +(site3.Period15 ?? 0) + projectList3[i].Period15;
                site3.Period16 = +(site3.Period16 ?? 0) + projectList3[i].Period16;

                site1_0.Period1 = +(site1_0.Period1 ?? 0) + projectList1_0[i].Period1;
                site1_0.Period2 = +(site1_0.Period2 ?? 0) + projectList1_0[i].Period2;
                site1_0.Period3 = +(site1_0.Period3 ?? 0) + projectList1_0[i].Period3;
                site1_0.Period4 = +(site1_0.Period4 ?? 0) + projectList1_0[i].Period4;
                site1_0.Period5 = +(site1_0.Period5 ?? 0) + projectList1_0[i].Period5;
                site1_0.Period6 = +(site1_0.Period6 ?? 0) + projectList1_0[i].Period6;
                site1_0.Period7 = +(site1_0.Period7 ?? 0) + projectList1_0[i].Period7;
                site1_0.Period8 = +(site1_0.Period8 ?? 0) + projectList1_0[i].Period8;
                site1_0.Period9 = +(site1_0.Period9 ?? 0) + projectList1_0[i].Period9;
                site1_0.Period10 = +(site1_0.Period10 ?? 0) + projectList1_0[i].Period10;
                site1_0.Period11 = +(site1_0.Period11 ?? 0) + projectList1_0[i].Period11;
                site1_0.Period12 = +(site1_0.Period12 ?? 0) + projectList1_0[i].Period12;
                site1_0.Period13 = +(site1_0.Period13 ?? 0) + projectList1_0[i].Period13;
                site1_0.Period14 = +(site1_0.Period14 ?? 0) + projectList1_0[i].Period14;
                site1_0.Period15 = +(site1_0.Period15 ?? 0) + projectList1_0[i].Period15;
                site1_0.Period16 = +(site1_0.Period16 ?? 0) + projectList1_0[i].Period16;

                site2_1.Period1 = +(site2_1.Period1 ?? 0) + projectList2_1[i].Period1;
                site2_1.Period2 = +(site2_1.Period2 ?? 0) + projectList2_1[i].Period2;
                site2_1.Period3 = +(site2_1.Period3 ?? 0) + projectList2_1[i].Period3;
                site2_1.Period4 = +(site2_1.Period4 ?? 0) + projectList2_1[i].Period4;
                site2_1.Period5 = +(site2_1.Period5 ?? 0) + projectList2_1[i].Period5;
                site2_1.Period6 = +(site2_1.Period6 ?? 0) + projectList2_1[i].Period6;
                site2_1.Period7 = +(site2_1.Period7 ?? 0) + projectList2_1[i].Period7;
                site2_1.Period8 = +(site2_1.Period8 ?? 0) + projectList2_1[i].Period8;
                site2_1.Period9 = +(site2_1.Period9 ?? 0) + projectList2_1[i].Period9;
                site2_1.Period10 = +(site2_1.Period10 ?? 0) + projectList2_1[i].Period10;
                site2_1.Period11 = +(site2_1.Period11 ?? 0) + projectList2_1[i].Period11;
                site2_1.Period12 = +(site2_1.Period12 ?? 0) + projectList2_1[i].Period12;
                site2_1.Period13 = +(site2_1.Period13 ?? 0) + projectList2_1[i].Period13;
                site2_1.Period14 = +(site2_1.Period14 ?? 0) + projectList2_1[i].Period14;
                site2_1.Period15 = +(site2_1.Period15 ?? 0) + projectList2_1[i].Period15;
                site2_1.Period16 = +(site2_1.Period16 ?? 0) + projectList2_1[i].Period16;
            }

            sitePercent2_1.Period1 = (site1_0.Period1 == 0) ? 0 : (site2_1.Period1 / site1_0.Period1);
            sitePercent2_1.Period2 = (site1_0.Period2 == 0) ? 0 : (site2_1.Period2 / site1_0.Period2);
            sitePercent2_1.Period3 = (site1_0.Period3 == 0) ? 0 : (site2_1.Period3 / site1_0.Period3);
            sitePercent2_1.Period4 = (site1_0.Period4 == 0) ? 0 : (site2_1.Period5 / site1_0.Period4);
            sitePercent2_1.Period5 = (site1_0.Period5 == 0) ? 0 : (site2_1.Period5 / site1_0.Period5);
            sitePercent2_1.Period6 = (site1_0.Period6 == 0) ? 0 : (site2_1.Period6 / site1_0.Period6);
            sitePercent2_1.Period7 = (site1_0.Period7 == 0) ? 0 : (site2_1.Period7 / site1_0.Period7);
            sitePercent2_1.Period8 = (site1_0.Period8 == 0) ? 0 : (site2_1.Period8 / site1_0.Period8);
            sitePercent2_1.Period9 = (site1_0.Period9 == 0) ? 0 : (site2_1.Period9 / site1_0.Period9);
            sitePercent2_1.Period10 = (site1_0.Period10 == 0) ? 0 : (site2_1.Period10 / site1_0.Period10);
            sitePercent2_1.Period11 = (site1_0.Period11 == 0) ? 0 : (site2_1.Period11 / site1_0.Period11);
            sitePercent2_1.Period12 = (site1_0.Period12 == 0) ? 0 : (site2_1.Period12 / site1_0.Period12);
            sitePercent2_1.Period13 = (site1_0.Period13 == 0) ? 0 : (site2_1.Period13 / site1_0.Period13);
            sitePercent2_1.Period14 = (site1_0.Period14 == 0) ? 0 : (site2_1.Period14 / site1_0.Period14);
            sitePercent2_1.Period15 = (site1_0.Period15 == 0) ? 0 : (site2_1.Period15 / site1_0.Period15);
            sitePercent2_1.Period16 = (site1_0.Period16 == 0) ? 0 : (site2_1.Period16 / site1_0.Period16);
            #endregion

            //********计算percent= {Project (1.1 + 1.2 + 1.3) ÷ Site (1.1 + 1.2 + 1.3)+ [Project(1.1 + 1.2) @ Project(2.1 %) + Project(1.3)] ÷ [Site(1.1 + 1.2) @ Site(2.1 %) + Site(1.3)]} ÷ 2
            try
            {
                for (int i = 0; i < projectList1.Count; i++)
                {
                    FOL_Input_2_1 newItem_Result = new FOL_Input_2_1();
                    newItem_Result.CustomerBPCCode = projectList1[i].CustomerBPCCode;
                    newItem_Result.CustomerOutputCode = projectList1[i].CustomerOutputCode;
                    newItem_Result.DIM1 = projectList1[i].DIM1;
                    newItem_Result.Order = projectList1[i].Order;
                    newItem_Result.BU = projectList1[i].BU;
                    newItem_Result.Segment = projectList1[i].Segment;
                    #region
                    if ((site1.Period1 + site2.Period1 + site3.Period1) == 0 || ((site1.Period1 + site2.Period1) * sitePercent2_1.Period1 + site3.Period1) == 0)
                        newItem_Result.Period1 = 0;
                    else
                        newItem_Result.Period1 = ((projectList1[i].Period1 + projectList2[i].Period1 + projectList3[i].Period1) / (site1.Period1 + site2.Period1 + site3.Period1) + ((projectList1[i].Period1 + projectList2[i].Period1) * percentList2_1[i].Period1 + projectList3[i].Period1) / ((site1.Period1 + site2.Period1) * sitePercent2_1.Period1 + site3.Period1)) / 2;
                    if ((site1.Period2 + site2.Period2 + site3.Period2) == 0 || ((site1.Period2 + site2.Period2) * sitePercent2_1.Period2 + site3.Period2) == 0)
                        newItem_Result.Period2 = 0;
                    else
                        newItem_Result.Period2 = ((projectList1[i].Period2 + projectList2[i].Period2 + projectList3[i].Period2) / (site1.Period2 + site2.Period2 + site3.Period2) + ((projectList1[i].Period2 + projectList2[i].Period2) * percentList2_1[i].Period2 + projectList3[i].Period2) / ((site1.Period2 + site2.Period2) * sitePercent2_1.Period2 + site3.Period2)) / 2;
                    if ((site1.Period3 + site2.Period3 + site3.Period3) == 0 || ((site1.Period3 + site2.Period3) * sitePercent2_1.Period3 + site3.Period3) == 0)
                        newItem_Result.Period3 = 0;
                    else
                        newItem_Result.Period3 = ((projectList1[i].Period3 + projectList2[i].Period3 + projectList3[i].Period3) / (site1.Period3 + site2.Period3 + site3.Period3) + ((projectList1[i].Period3 + projectList2[i].Period3) * percentList2_1[i].Period3 + projectList3[i].Period3) / ((site1.Period3 + site2.Period3) * sitePercent2_1.Period3 + site3.Period3)) / 2;
                    if ((site1.Period4 + site2.Period4 + site3.Period4) == 0 || ((site1.Period4 + site2.Period4) * sitePercent2_1.Period4 + site3.Period4) == 0)
                        newItem_Result.Period4 = 0;
                    else
                        newItem_Result.Period4 = ((projectList1[i].Period4 + projectList2[i].Period4 + projectList3[i].Period4) / (site1.Period4 + site2.Period4 + site3.Period4) + ((projectList1[i].Period4 + projectList2[i].Period4) * percentList2_1[i].Period4 + projectList3[i].Period4) / ((site1.Period4 + site2.Period4) * sitePercent2_1.Period4 + site3.Period4)) / 2;
                    if ((site1.Period5 + site2.Period5 + site3.Period5) == 0 || ((site1.Period5 + site2.Period5) * sitePercent2_1.Period5 + site3.Period5) == 0)
                        newItem_Result.Period5 = 0;
                    else
                        newItem_Result.Period5 = ((projectList1[i].Period5 + projectList2[i].Period5 + projectList3[i].Period5) / (site1.Period5 + site2.Period5 + site3.Period5) + ((projectList1[i].Period5 + projectList2[i].Period5) * percentList2_1[i].Period5 + projectList3[i].Period5) / ((site1.Period5 + site2.Period5) * sitePercent2_1.Period5 + site3.Period5)) / 2;
                    if ((site1.Period6 + site2.Period6 + site3.Period6) == 0 || ((site1.Period6 + site2.Period6) * sitePercent2_1.Period6 + site3.Period6) == 0)
                        newItem_Result.Period6 = 0;
                    else
                        newItem_Result.Period6 = ((projectList1[i].Period6 + projectList2[i].Period6 + projectList3[i].Period6) / (site1.Period6 + site2.Period6 + site3.Period6) + ((projectList1[i].Period6 + projectList2[i].Period6) * percentList2_1[i].Period6 + projectList3[i].Period6) / ((site1.Period6 + site2.Period6) * sitePercent2_1.Period6 + site3.Period6)) / 2;
                    if ((site1.Period7 + site2.Period7 + site3.Period7) == 0 || ((site1.Period7 + site2.Period7) * sitePercent2_1.Period7 + site3.Period7) == 0)
                        newItem_Result.Period7 = 0;
                    else
                        newItem_Result.Period7 = ((projectList1[i].Period7 + projectList2[i].Period7 + projectList3[i].Period7) / (site1.Period7 + site2.Period7 + site3.Period7) + ((projectList1[i].Period7 + projectList2[i].Period7) * percentList2_1[i].Period7 + projectList3[i].Period7) / ((site1.Period7 + site2.Period7) * sitePercent2_1.Period7 + site3.Period7)) / 2;
                    if ((site1.Period8 + site2.Period8 + site3.Period8) == 0 || ((site1.Period8 + site2.Period8) * sitePercent2_1.Period8 + site3.Period8) == 0)
                        newItem_Result.Period8 = 0;
                    else
                        newItem_Result.Period8 = ((projectList1[i].Period8 + projectList2[i].Period8 + projectList3[i].Period8) / (site1.Period8 + site2.Period8 + site3.Period8) + ((projectList1[i].Period8 + projectList2[i].Period8) * percentList2_1[i].Period8 + projectList3[i].Period8) / ((site1.Period8 + site2.Period8) * sitePercent2_1.Period8 + site3.Period8)) / 2;
                    if ((site1.Period9 + site2.Period9 + site3.Period9) == 0 || ((site1.Period9 + site2.Period9) * sitePercent2_1.Period9 + site3.Period9) == 0)
                        newItem_Result.Period9 = 0;
                    else
                        newItem_Result.Period9 = ((projectList1[i].Period9 + projectList2[i].Period9 + projectList3[i].Period9) / (site1.Period9 + site2.Period9 + site3.Period9) + ((projectList1[i].Period9 + projectList2[i].Period9) * percentList2_1[i].Period9 + projectList3[i].Period9) / ((site1.Period9 + site2.Period9) * sitePercent2_1.Period9 + site3.Period9)) / 2;
                    if ((site1.Period10 + site2.Period10 + site3.Period10) == 0 || ((site1.Period10 + site2.Period10) * sitePercent2_1.Period10 + site3.Period10) == 0)
                        newItem_Result.Period10 = 0;
                    else
                        newItem_Result.Period10 = ((projectList1[i].Period10 + projectList2[i].Period10 + projectList3[i].Period10) / (site1.Period10 + site2.Period10 + site3.Period10) + ((projectList1[i].Period10 + projectList2[i].Period10) * percentList2_1[i].Period10 + projectList3[i].Period10) / ((site1.Period10 + site2.Period10) * sitePercent2_1.Period10 + site3.Period10)) / 2;
                    if ((site1.Period11 + site2.Period11 + site3.Period11) == 0 || ((site1.Period11 + site2.Period11) * sitePercent2_1.Period11 + site3.Period11) == 0)
                        newItem_Result.Period11 = 0;
                    else
                        newItem_Result.Period11 = ((projectList1[i].Period11 + projectList2[i].Period11 + projectList3[i].Period11) / (site1.Period11 + site2.Period11 + site3.Period11) + ((projectList1[i].Period11 + projectList2[i].Period11) * percentList2_1[i].Period11 + projectList3[i].Period11) / ((site1.Period11 + site2.Period11) * sitePercent2_1.Period11 + site3.Period11)) / 2;
                    if ((site1.Period12 + site2.Period12 + site3.Period12) == 0 || ((site1.Period12 + site2.Period12) * sitePercent2_1.Period12 + site3.Period12) == 0)
                        newItem_Result.Period12 = 0;
                    else
                        newItem_Result.Period12 = ((projectList1[i].Period12 + projectList2[i].Period12 + projectList3[i].Period12) / (site1.Period12 + site2.Period12 + site3.Period12) + ((projectList1[i].Period12 + projectList2[i].Period12) * percentList2_1[i].Period12 + projectList3[i].Period12) / ((site1.Period12 + site2.Period12) * sitePercent2_1.Period12 + site3.Period12)) / 2;
                    if ((site1.Period13 + site2.Period13 + site3.Period13) == 0 || ((site1.Period13 + site2.Period13) * sitePercent2_1.Period13 + site3.Period13) == 0)
                        newItem_Result.Period13 = 0;
                    else
                        newItem_Result.Period13 = ((projectList1[i].Period13 + projectList2[i].Period13 + projectList3[i].Period13) / (site1.Period13 + site2.Period13 + site3.Period13) + ((projectList1[i].Period13 + projectList2[i].Period13) * percentList2_1[i].Period13 + projectList3[i].Period13) / ((site1.Period13 + site2.Period13) * sitePercent2_1.Period13 + site3.Period13)) / 2;
                    if ((site1.Period14 + site2.Period14 + site3.Period14) == 0 || ((site1.Period14 + site2.Period14) * sitePercent2_1.Period14 + site3.Period14) == 0)
                        newItem_Result.Period14 = 0;
                    else
                        newItem_Result.Period14 = ((projectList1[i].Period14 + projectList2[i].Period14 + projectList3[i].Period14) / (site1.Period14 + site2.Period14 + site3.Period14) + ((projectList1[i].Period14 + projectList2[i].Period14) * percentList2_1[i].Period14 + projectList3[i].Period14) / ((site1.Period14 + site2.Period14) * sitePercent2_1.Period14 + site3.Period14)) / 2;
                    if ((site1.Period15 + site2.Period15 + site3.Period15) == 0 || ((site1.Period15 + site2.Period15) * sitePercent2_1.Period15 + site3.Period15) == 0)
                        newItem_Result.Period15 = 0;
                    else
                        newItem_Result.Period15 = ((projectList1[i].Period15 + projectList2[i].Period15 + projectList3[i].Period15) / (site1.Period15 + site2.Period15 + site3.Period15) + ((projectList1[i].Period15 + projectList2[i].Period15) * percentList2_1[i].Period15 + projectList3[i].Period15) / ((site1.Period15 + site2.Period15) * sitePercent2_1.Period15 + site3.Period15)) / 2;
                    if ((site1.Period16 + site2.Period16 + site3.Period16) == 0 || ((site1.Period16 + site2.Period16) * sitePercent2_1.Period16 + site3.Period16) == 0)
                        newItem_Result.Period16 = 0;
                    else
                        newItem_Result.Period16 = ((projectList1[i].Period16 + projectList2[i].Period16 + projectList3[i].Period16) / (site1.Period16 + site2.Period16 + site3.Period16) + ((projectList1[i].Period16 + projectList2[i].Period16) * percentList2_1[i].Period16 + projectList3[i].Period16) / ((site1.Period16 + site2.Period16) * sitePercent2_1.Period16 + site3.Period16)) / 2;
                    #endregion
                    newItem_Result.Type = type;
                    newItem_Result.Version = version;
                    newItem_Result.InsertUser = "";
                    newItem_Result.InsertDate = System.DateTime.Now;
                    newItem_Result.Department = "";
                    listResult.Add(newItem_Result);
                }
                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList());
                dc.FOL_Input_2_1.InsertAllOnSubmit(listResult);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        //8.0
        public string Calculate8_0(string version, string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list4 = dc.FOL_Input_1_1.Where(x => x.Type == "1.4 Rev Recog - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_1_1> list5 = dc.FOL_Input_1_1.Where(x => x.Type == "1.5 Rev Reversal - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_2_1> list6 = dc.FOL_Input_2_1.Where(x => x.Type == "8.0 Corp. Alloc. (%)" && x.Version == version).ToList();
            if (list1.Count == 0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count == 0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count == 0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list4.Count == 0)
            { message = message + "该月份的1.4 Rev Recog - OT Contract数据为空;"; return message; }
            if (list5.Count == 0)
            { message = message + "该月份的1.5 Rev Reversal - OT Contract数据为空;"; return message; }
            if (list6.Count == 0)
            { message = message + "该月份的8.0 Corp. Alloc. (%)的百分比数据为空;"; return message; }

            List<FOL_Input_1_1> list = new List<FOL_Input_1_1>();
            try
            {
                int i = 0;
                foreach (FOL_Input_1_1 item in list1)
                {
                    i++;
                    FOL_Input_1_1 newItem = new FOL_Input_1_1
                    {
                        GLBPCCode = item.GLBPCCode,
                        GLBPCDescription = item.GLBPCDescription,
                        GLOutputCode = item.GLOutputCode,
                        CustomerBPCCode = item.CustomerBPCCode,
                        DIM1 = item.DIM1,
                        Order = item.Order,
                        CustomerOutputCode = item.CustomerOutputCode,
                        BU = item.BU,
                        Segment = item.Segment,
                        Period1 = (item.Period1 + list2[i].Period1 + list3[i].Period1 + list4[i].Period1 + list5[i].Period1) * list6[i].Period1,
                        Period2 = (item.Period2 + list2[i].Period2 + list3[i].Period2 + list4[i].Period2 + list5[i].Period2) * list6[i].Period2,
                        Period3 = (item.Period3 + list2[i].Period3 + list3[i].Period3 + list4[i].Period3 + list5[i].Period3) * list6[i].Period3,
                        Period4 = (item.Period4 + list2[i].Period4 + list3[i].Period4 + list4[i].Period4 + list5[i].Period4) * list6[i].Period4,
                        Period5 = (item.Period5 + list2[i].Period5 + list3[i].Period5 + list4[i].Period5 + list5[i].Period5) * list6[i].Period5,
                        Period6 = (item.Period6 + list2[i].Period6 + list3[i].Period6 + list4[i].Period6 + list5[i].Period6) * list6[i].Period6,
                        Period7 = (item.Period7 + list2[i].Period7 + list3[i].Period7 + list4[i].Period7 + list5[i].Period7) * list6[i].Period7,
                        Period8 = (item.Period8 + list2[i].Period8 + list3[i].Period8 + list4[i].Period8 + list5[i].Period8) * list6[i].Period8,
                        Period9 = (item.Period9 + list2[i].Period9 + list3[i].Period9 + list4[i].Period9 + list5[i].Period9) * list6[i].Period9,
                        Period10 = (item.Period10 + list2[i].Period10 + list3[i].Period10 + list4[i].Period10 + list5[i].Period10) * list6[i].Period10,
                        Period11 = (item.Period11 + list2[i].Period11 + list3[i].Period11 + list4[i].Period11 + list5[i].Period11) * list6[i].Period11,
                        Period12 = (item.Period12 + list2[i].Period12 + list3[i].Period12 + list4[i].Period12 + list5[i].Period12) * list6[i].Period12,
                        Period13 = (item.Period13 + list2[i].Period13 + list3[i].Period13 + list4[i].Period13 + list5[i].Period13) * list6[i].Period13,
                        Period14 = (item.Period14 + list2[i].Period14 + list3[i].Period14 + list4[i].Period14 + list5[i].Period14) * list6[i].Period14,
                        Period15 = (item.Period15 + list2[i].Period15 + list3[i].Period15 + list4[i].Period15 + list5[i].Period15) * list6[i].Period15,
                        Type = type,
                        InsertDate = System.DateTime.Now,
                        InsertUser = "",
                        Version = version,
                    };
                }
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == "8.0 Corp. Alloc. (%)").ToList());
                dc.FOL_Input_1_1.InsertAllOnSubmit(list);
                dc.SubmitChanges();
                message = "Success";
            }
            catch (Exception ex)
            {
                message = message + ex.Message;
            }
            return message;
        }

        public string getType(string type)
        {
            string actualType = "";
            switch (type)
            {
                case "1.0":
                    actualType = "1.0 Total Sales($)";
                    break;
                case "1.1":
                    actualType = "1.1 Gross Sales - External($)";
                    break;
                case "1.2":
                    actualType = "1.2 Gross Sales - Interco($)";
                    break;
                case "1.3":
                    actualType = "1.3 Gross Sales - Recoveries($)";
                    break;
                case "1.4":
                    actualType = "1.4 Rev Recog - OT Contract";
                    break;
                case "1.5":
                    actualType = "1.5 Rev Reversal - OT Contract";
                    break;
                case "2.0":
                    actualType = "2.0 MCOS($)";
                    break;
                case "2.1":
                    actualType = "2.1 Std VAM % from Ops(%)";
                    break;
                case "2.2":
                    actualType = "2.2 MCOS Recoveries($)";
                    break;
                case "2.3":
                    actualType = "2.3 COGS Recog - OT Contract";
                    break;
                case "2.4":
                    actualType = "2.4 COGS Reversal - OT Contract";
                    break;
                case "3.0":
                    actualType = "3.0 Material Margin ($)";
                    break;
                case "3.1":
                    actualType = "3.1 PPV(%&$)";
                    break;
                case "3.2":
                    actualType = "3.2 FCP-PPV($)";
                    break;
                case "3.3":
                    actualType = "3.3 Alloc-PPV($)";
                    break;
                case "3.4":
                    actualType = "3.4 FCP_Alloc BPC Working";
                    break;
                case "4.1":
                    actualType = "4.1 Material Loss (%)";
                    break;
                case "4.2":
                    actualType = "4.2 Freight in (%)";
                    break;
                case "4.3":
                    actualType = "4.3 Freight out (%)";
                    break;
                case "4.4":
                    actualType = "4.4 EDM (%)";
                    break;
                case "4.5":
                    actualType = "4.5 MCOSO (%)";
                    break;
                case "4.6":
                    actualType = "4.6 Subcontract ($)";
                    break;
                case "4.7":
                    actualType = "4.7 Inv. Reserve ($)";
                    break;
                case "5.1":
                    actualType = "5.1 DL (%)";
                    break;
                case "6.0":
                    actualType = "6.0 Total MOH($)";
                    break;
                case "6.1":
                    actualType = "6.1 IDL (%&#)";
                    break;
                case "6.2":
                    actualType = "6.2 Depn (%)";
                    break;
                case "6.3":
                    actualType = "6.3 FUS. (%)";
                    break;
                case "6.4":
                    actualType = "6.4 MOHO (%)";
                    break;
                case "6.5":
                    actualType = "6.5 OEUP (%)";
                    break;
                case "6.6":
                    actualType = "6.6 ELEC (%)";
                    break;
                case "7.0":
                    actualType = "7.0 Total SG&A($)";
                    break;
                case "7.1":
                    actualType = "7.1 G.A.(%)";
                    break;
                case "7.2":
                    actualType = "7.2 S.M.(%)";
                    break;
                case "8.0":
                    actualType = "8.0 Corp. Alloc. (%)";
                    break;
                case "9.0":
                    actualType = "9.0 Other Costs ($)";
                    break;
                case "10.0":
                    actualType = "10.0 Change of CLOH ($)";
                    break;
                case "10.1":
                    actualType = "10.1 CLOH Recog OT ($)";
                    break;
                case "10.2":
                    actualType = "10.2 CLOH Reversal OT ($)";
                    break;
            }
            return actualType;
        }
    }
}