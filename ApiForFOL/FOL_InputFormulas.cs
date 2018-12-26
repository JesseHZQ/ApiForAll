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
            if (list1.Count==0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list4.Count==0)
            { message = message + "该月份的1.4 Rev Recog - OT Contract数据为空;"; return message; }
            if (list5.Count==0)
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
            if (list1.Count==0)
            { message = message + "该月份的2.1 Std VAM % from Ops(%)数据为空"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的2.2 MCOS Recoveries($)数据为空"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的2.3 COGS Recog - OT Contract数据为空"; return message; }
            if (list4.Count==0)
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
            if (list1.Count==0)
            { message = message + "该月份的1.1 Gross Sales - External($)数据为空"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空"; return message; }
            if (list3.Count==0)
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
            if (list1.Count==0)
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
            if (list1.Count==0)
            { message = "该月份的3.1 PPV(%&$)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的3.2 FCP-PPV($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的3.3 Alloc-PPV($)数据为空;"; return message; }
            if (list_percent.Count==0)
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

            if (list1.Count==0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list_percent.Count==0)
            { message = message + "该月份的" + type + "百分比数据为空;"; return message; }
            if (list_amount.Count==0)
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
            if (list1.Count==0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list_percent.Count==0)
            { message = message + "该月份的4.1 Material Loss (%)百分比数据为空;"; return message; }
            if (list_amount.Count==0)
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
            if (list1.Count==0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的5.1 DL (%)百分比数据为空;"; return message; }
            if (list4.Count==0)
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

        //6.0
        public string Calculate6_0(string version,string type)
        {
            string message = "";
            List<FOL_Input_6_0_TotalMOH> list = new List<FOL_Input_6_0_TotalMOH>();
            List<string> list_department = dc.FOL_CCMapping.Where(x => x.Type == "MOH").Select(x => x.CCDescription).ToList();
            List<FOL_Input_6_0_SumByTypeModule> list_SumByTypeModule = dc.FOL_Input_6_0_SumByTypeModule.ToList();
            List<FOL_Input_6_0_SumByType> list_SumByType = new List<FOL_Input_6_0_SumByType>();
            try
            {
                for (int n = 0; n < list_SumByTypeModule.Count; n++)
                {
                    FOL_Input_6_0_SumByType newItem_SumByType = new FOL_Input_6_0_SumByType();
                    newItem_SumByType.GLBPCCode = list_SumByTypeModule[n].GLBPCCode;
                    newItem_SumByType.GLBPCDescription = list_SumByTypeModule[n].GLBPCDescription;
                    newItem_SumByType.GLOutputCode = list_SumByTypeModule[n].GLOutputCode;
                    newItem_SumByType.PARENTH1 = list_SumByTypeModule[n].PARENTH1;
                    newItem_SumByType.UploadCode = list_SumByTypeModule[n].UploadCode;
                    newItem_SumByType.BPCOutputCode = list_SumByTypeModule[n].BPCOutputCode;
                    newItem_SumByType.BU = list_SumByTypeModule[n].BU;
                    newItem_SumByType.Segment = list_SumByTypeModule[n].Segment;
                    newItem_SumByType.Period1 = 0;
                    newItem_SumByType.Period2 = 0;
                    newItem_SumByType.Period3 = 0;
                    newItem_SumByType.Period4 = 0;
                    newItem_SumByType.Period5 = 0;
                    newItem_SumByType.Period6 = 0;
                    newItem_SumByType.Period7 = 0;
                    newItem_SumByType.Period8 = 0;
                    newItem_SumByType.Period9 = 0;
                    newItem_SumByType.Period10 = 0;
                    newItem_SumByType.Period11 = 0;
                    newItem_SumByType.Period12 = 0;
                    newItem_SumByType.Period13 = 0;
                    newItem_SumByType.Period14 = 0;
                    newItem_SumByType.Period15 = 0;
                    newItem_SumByType.Period16 = 0;
                    newItem_SumByType.Type = type;
                    newItem_SumByType.Version = version;
                    newItem_SumByType.InsertDate = System.DateTime.Now;
                    newItem_SumByType.InsertUser = "";
                    list_SumByType.Add(newItem_SumByType);
                }
                for (int i = 0; i < list_department.Count; i++)
                {
                    string department = list_department[i];
                    List<FOL_Input_6_0_TotalMOH> list_new = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department.Contains(department) && x.IsSubtotal != 11).ToList();
                    List<FOL_Input_6_0_TotalMOH> list_byDepAndBPCCode = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department.Contains(department) && x.IsSubtotal == 11).ToList();

                    FOL_Input_6_0_TotalMOH newItem = new FOL_Input_6_0_TotalMOH();

                    for (int j = 0; j < list_new.Count; j++)
                    {
                        //计算总和By BPCCode By Dep
                        #region
                        if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[0].PARENTH1)
                        {
                            list_byDepAndBPCCode[0].Period1 = (list_byDepAndBPCCode[0].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[0].Period2 = (list_byDepAndBPCCode[0].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[0].Period3 = (list_byDepAndBPCCode[0].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[0].Period4 = (list_byDepAndBPCCode[0].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[0].Period5 = (list_byDepAndBPCCode[0].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[0].Period6 = (list_byDepAndBPCCode[0].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[0].Period7 = (list_byDepAndBPCCode[0].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[0].Period8 = (list_byDepAndBPCCode[0].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[0].Period9 = (list_byDepAndBPCCode[0].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[0].Period10 = (list_byDepAndBPCCode[0].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[0].Period11 = (list_byDepAndBPCCode[0].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[0].Period12 = (list_byDepAndBPCCode[0].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[0].Period13 = (list_byDepAndBPCCode[0].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[0].Period14 = (list_byDepAndBPCCode[0].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[0].Period15 = (list_byDepAndBPCCode[0].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[0].Period16 = (list_byDepAndBPCCode[0].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[0].Period1 = (list_SumByType[0].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[0].Period2 = (list_SumByType[0].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[0].Period3 = (list_SumByType[0].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[0].Period4 = (list_SumByType[0].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[0].Period5 = (list_SumByType[0].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[0].Period6 = (list_SumByType[0].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[0].Period7 = (list_SumByType[0].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[0].Period8 = (list_SumByType[0].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[0].Period9 = (list_SumByType[0].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[0].Period10 = (list_SumByType[0].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[0].Period11 = (list_SumByType[0].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[0].Period12 = (list_SumByType[0].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[0].Period13 = (list_SumByType[0].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[0].Period14 = (list_SumByType[0].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[0].Period15 = (list_SumByType[0].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[0].Period16 = (list_SumByType[0].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[1].PARENTH1)
                        {
                            list_byDepAndBPCCode[1].Period1 = (list_byDepAndBPCCode[1].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[1].Period2 = (list_byDepAndBPCCode[1].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[1].Period3 = (list_byDepAndBPCCode[1].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[1].Period4 = (list_byDepAndBPCCode[1].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[1].Period5 = (list_byDepAndBPCCode[1].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[1].Period6 = (list_byDepAndBPCCode[1].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[1].Period7 = (list_byDepAndBPCCode[1].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[1].Period8 = (list_byDepAndBPCCode[1].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[1].Period9 = (list_byDepAndBPCCode[1].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[1].Period10 = (list_byDepAndBPCCode[1].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[1].Period11 = (list_byDepAndBPCCode[1].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[1].Period12 = (list_byDepAndBPCCode[1].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[1].Period13 = (list_byDepAndBPCCode[1].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[1].Period14 = (list_byDepAndBPCCode[1].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[1].Period15 = (list_byDepAndBPCCode[1].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[1].Period16 = (list_byDepAndBPCCode[1].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[1].Period1 = (list_SumByType[1].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[1].Period2 = (list_SumByType[1].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[1].Period3 = (list_SumByType[1].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[1].Period4 = (list_SumByType[1].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[1].Period5 = (list_SumByType[1].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[1].Period6 = (list_SumByType[1].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[1].Period7 = (list_SumByType[1].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[1].Period8 = (list_SumByType[1].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[1].Period9 = (list_SumByType[1].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[1].Period10 = (list_SumByType[1].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[1].Period11 = (list_SumByType[1].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[1].Period12 = (list_SumByType[1].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[1].Period13 = (list_SumByType[1].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[1].Period14 = (list_SumByType[1].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[1].Period15 = (list_SumByType[1].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[1].Period16 = (list_SumByType[1].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[2].PARENTH1)
                        {
                            list_byDepAndBPCCode[2].Period1 = (list_byDepAndBPCCode[2].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[2].Period2 = (list_byDepAndBPCCode[2].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[2].Period3 = (list_byDepAndBPCCode[2].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[2].Period4 = (list_byDepAndBPCCode[2].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[2].Period5 = (list_byDepAndBPCCode[2].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[2].Period6 = (list_byDepAndBPCCode[2].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[2].Period7 = (list_byDepAndBPCCode[2].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[2].Period8 = (list_byDepAndBPCCode[2].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[2].Period9 = (list_byDepAndBPCCode[2].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[2].Period10 = (list_byDepAndBPCCode[2].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[2].Period11 = (list_byDepAndBPCCode[2].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[2].Period12 = (list_byDepAndBPCCode[2].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[2].Period13 = (list_byDepAndBPCCode[2].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[2].Period14 = (list_byDepAndBPCCode[2].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[2].Period15 = (list_byDepAndBPCCode[2].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[2].Period16 = (list_byDepAndBPCCode[2].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[2].Period1 = (list_SumByType[2].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[2].Period2 = (list_SumByType[2].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[2].Period3 = (list_SumByType[2].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[2].Period4 = (list_SumByType[2].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[2].Period5 = (list_SumByType[2].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[2].Period6 = (list_SumByType[2].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[2].Period7 = (list_SumByType[2].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[2].Period8 = (list_SumByType[2].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[2].Period9 = (list_SumByType[2].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[2].Period10 = (list_SumByType[2].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[2].Period11 = (list_SumByType[2].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[2].Period12 = (list_SumByType[2].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[2].Period13 = (list_SumByType[2].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[2].Period14 = (list_SumByType[2].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[2].Period15 = (list_SumByType[2].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[2].Period16 = (list_SumByType[2].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[3].PARENTH1)
                        {
                            list_byDepAndBPCCode[3].Period1 = (list_byDepAndBPCCode[3].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[3].Period2 = (list_byDepAndBPCCode[3].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[3].Period3 = (list_byDepAndBPCCode[3].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[3].Period4 = (list_byDepAndBPCCode[3].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[3].Period5 = (list_byDepAndBPCCode[3].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[3].Period6 = (list_byDepAndBPCCode[3].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[3].Period7 = (list_byDepAndBPCCode[3].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[3].Period8 = (list_byDepAndBPCCode[3].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[3].Period9 = (list_byDepAndBPCCode[3].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[3].Period10 = (list_byDepAndBPCCode[3].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[3].Period11 = (list_byDepAndBPCCode[3].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[3].Period12 = (list_byDepAndBPCCode[3].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[3].Period13 = (list_byDepAndBPCCode[3].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[3].Period14 = (list_byDepAndBPCCode[3].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[3].Period15 = (list_byDepAndBPCCode[3].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[3].Period16 = (list_byDepAndBPCCode[3].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[3].Period1 = (list_SumByType[3].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[3].Period2 = (list_SumByType[3].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[3].Period3 = (list_SumByType[3].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[3].Period4 = (list_SumByType[3].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[3].Period5 = (list_SumByType[3].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[3].Period6 = (list_SumByType[3].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[3].Period7 = (list_SumByType[3].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[3].Period8 = (list_SumByType[3].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[3].Period9 = (list_SumByType[3].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[3].Period10 = (list_SumByType[3].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[3].Period11 = (list_SumByType[3].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[3].Period12 = (list_SumByType[3].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[3].Period13 = (list_SumByType[3].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[3].Period14 = (list_SumByType[3].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[3].Period15 = (list_SumByType[3].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[3].Period16 = (list_SumByType[3].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[4].PARENTH1)
                        {
                            list_byDepAndBPCCode[4].Period1 = (list_byDepAndBPCCode[4].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[4].Period2 = (list_byDepAndBPCCode[4].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[4].Period3 = (list_byDepAndBPCCode[4].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[4].Period4 = (list_byDepAndBPCCode[4].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[4].Period5 = (list_byDepAndBPCCode[4].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[4].Period6 = (list_byDepAndBPCCode[4].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[4].Period7 = (list_byDepAndBPCCode[4].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[4].Period8 = (list_byDepAndBPCCode[4].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[4].Period9 = (list_byDepAndBPCCode[4].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[4].Period10 = (list_byDepAndBPCCode[4].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[4].Period11 = (list_byDepAndBPCCode[4].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[4].Period12 = (list_byDepAndBPCCode[4].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[4].Period13 = (list_byDepAndBPCCode[4].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[4].Period14 = (list_byDepAndBPCCode[4].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[4].Period15 = (list_byDepAndBPCCode[4].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[4].Period16 = (list_byDepAndBPCCode[4].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[4].Period1 = (list_SumByType[4].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[4].Period2 = (list_SumByType[4].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[4].Period3 = (list_SumByType[4].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[4].Period4 = (list_SumByType[4].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[4].Period5 = (list_SumByType[4].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[4].Period6 = (list_SumByType[4].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[4].Period7 = (list_SumByType[4].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[4].Period8 = (list_SumByType[4].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[4].Period9 = (list_SumByType[4].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[4].Period10 = (list_SumByType[4].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[4].Period11 = (list_SumByType[4].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[4].Period12 = (list_SumByType[4].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[4].Period13 = (list_SumByType[4].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[4].Period14 = (list_SumByType[4].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[4].Period15 = (list_SumByType[4].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[4].Period16 = (list_SumByType[1].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[5].PARENTH1)
                        {
                            list_byDepAndBPCCode[5].Period1 = (list_byDepAndBPCCode[5].Period1 ?? 0) + list_new[j].Period1;
                            list_byDepAndBPCCode[5].Period2 = (list_byDepAndBPCCode[5].Period2 ?? 0) + list_new[j].Period2;
                            list_byDepAndBPCCode[5].Period3 = (list_byDepAndBPCCode[5].Period3 ?? 0) + list_new[j].Period3;
                            list_byDepAndBPCCode[5].Period4 = (list_byDepAndBPCCode[5].Period4 ?? 0) + list_new[j].Period4;
                            list_byDepAndBPCCode[5].Period5 = (list_byDepAndBPCCode[5].Period5 ?? 0) + list_new[j].Period5;
                            list_byDepAndBPCCode[5].Period6 = (list_byDepAndBPCCode[5].Period6 ?? 0) + list_new[j].Period6;
                            list_byDepAndBPCCode[5].Period7 = (list_byDepAndBPCCode[5].Period7 ?? 0) + list_new[j].Period7;
                            list_byDepAndBPCCode[5].Period8 = (list_byDepAndBPCCode[5].Period8 ?? 0) + list_new[j].Period8;
                            list_byDepAndBPCCode[5].Period9 = (list_byDepAndBPCCode[5].Period9 ?? 0) + list_new[j].Period9;
                            list_byDepAndBPCCode[5].Period10 = (list_byDepAndBPCCode[5].Period10 ?? 0) + list_new[j].Period10;
                            list_byDepAndBPCCode[5].Period11 = (list_byDepAndBPCCode[5].Period11 ?? 0) + list_new[j].Period11;
                            list_byDepAndBPCCode[5].Period12 = (list_byDepAndBPCCode[5].Period12 ?? 0) + list_new[j].Period12;
                            list_byDepAndBPCCode[5].Period13 = (list_byDepAndBPCCode[5].Period13 ?? 0) + list_new[j].Period13;
                            list_byDepAndBPCCode[5].Period14 = (list_byDepAndBPCCode[5].Period14 ?? 0) + list_new[j].Period14;
                            list_byDepAndBPCCode[5].Period15 = (list_byDepAndBPCCode[5].Period15 ?? 0) + list_new[j].Period15;
                            list_byDepAndBPCCode[5].Period16 = (list_byDepAndBPCCode[5].Period16 ?? 0) + list_new[j].Period16;

                            list_SumByType[5].Period1 = (list_SumByType[5].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[5].Period2 = (list_SumByType[5].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[5].Period3 = (list_SumByType[5].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[5].Period4 = (list_SumByType[5].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[5].Period5 = (list_SumByType[5].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[5].Period6 = (list_SumByType[5].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[5].Period7 = (list_SumByType[5].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[5].Period8 = (list_SumByType[5].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[5].Period9 = (list_SumByType[5].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[5].Period10 = (list_SumByType[5].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[5].Period11 = (list_SumByType[5].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[5].Period12 = (list_SumByType[5].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[5].Period13 = (list_SumByType[5].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[5].Period14 = (list_SumByType[5].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[5].Period15 = (list_SumByType[5].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[5].Period16 = (list_SumByType[5].Period16 ?? 0) + list_new[j].Period16;
                        }
                        #endregion
                        //计算IsSubtotal=1的类型的和
                        #region
                        if (list_new[j].IsSubtotal == 1)
                        {
                            list_new[j].Period1 = newItem.Period1;
                            list_new[j].Period2 = newItem.Period2;
                            list_new[j].Period3 = newItem.Period3;
                            list_new[j].Period4 = newItem.Period4;
                            list_new[j].Period5 = newItem.Period5;
                            list_new[j].Period6 = newItem.Period6;
                            list_new[j].Period7 = newItem.Period7;
                            list_new[j].Period8 = newItem.Period8;
                            list_new[j].Period9 = newItem.Period9;
                            list_new[j].Period10 = newItem.Period10;
                            list_new[j].Period11 = newItem.Period11;
                            list_new[j].Period12 = newItem.Period12;
                            list_new[j].Period13 = newItem.Period13;
                            list_new[j].Period14 = newItem.Period14;
                            list_new[j].Period15 = newItem.Period15;
                            list_new[j].Period16 = newItem.Period16;
                            newItem = new FOL_Input_6_0_TotalMOH();
                        }
                        else
                        {
                            newItem.Period1 = (newItem.Period1 ?? 0) + list_new[j].Period1;
                            newItem.Period2 = (newItem.Period2 ?? 0) + list_new[j].Period2;
                            newItem.Period3 = (newItem.Period3 ?? 0) + list_new[j].Period3;
                            newItem.Period4 = (newItem.Period4 ?? 0) + list_new[j].Period4;
                            newItem.Period5 = (newItem.Period5 ?? 0) + list_new[j].Period5;
                            newItem.Period6 = (newItem.Period6 ?? 0) + list_new[j].Period6;
                            newItem.Period7 = (newItem.Period7 ?? 0) + list_new[j].Period7;
                            newItem.Period8 = (newItem.Period8 ?? 0) + list_new[j].Period8;
                            newItem.Period9 = (newItem.Period9 ?? 0) + list_new[j].Period9;
                            newItem.Period10 = (newItem.Period10 ?? 0) + list_new[j].Period10;
                            newItem.Period11 = (newItem.Period11 ?? 0) + list_new[j].Period11;
                            newItem.Period12 = (newItem.Period12 ?? 0) + list_new[j].Period12;
                            newItem.Period13 = (newItem.Period13 ?? 0) + list_new[j].Period13;
                            newItem.Period14 = (newItem.Period14 ?? 0) + list_new[j].Period14;
                            newItem.Period15 = (newItem.Period15 ?? 0) + list_new[j].Period15;
                            newItem.Period16 = (newItem.Period16 ?? 0) + list_new[j].Period16;
                        }
                        #endregion

                        ////计算总和By BPCCode
                        //for(int m=0;m< list_SumByType.Count; m++)
                        //{
                        //    if(list_new[j].BPCOutputCode == list_byDepAndBPCCode[0].PARENTH1)
                        //}
                    }
                }
                dc.FOL_Input_6_0_SumByType.DeleteAllOnSubmit(dc.FOL_Input_6_0_SumByType.Where(x => x.Type == "6.0 Total MOH($)" && x.Version == version).ToList());
                dc.FOL_Input_6_0_SumByType.InsertAllOnSubmit(list_SumByType);
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
        public string Calculate6_1(string version,string type, List<FOL_Input_6_1> list1)
        {
            List<FOL_Input_2_1> list = new List<FOL_Input_2_1>();
            List<FOL_Input_1_1> list_Dollar = new List<FOL_Input_1_1>();
            Dictionary<int, FOL_Input_6_1> dic_totalPeopleAmount = new Dictionary<int, FOL_Input_6_1>();
            string message = "";
            try
            {
                List<int> flags = new List<int>();
                int flag = 0;
                foreach (FOL_Input_6_1 item in list1)
                {
                    if (item.CustomerOutputCode == "CC31 - Facility" || item.CustomerOutputCode == "CC33 - PE" || item.CustomerOutputCode == "CC34 - Quality" || item.CustomerOutputCode == "CC37 - Security" || item.CustomerOutputCode == "CC42 - Planning" || item.CustomerOutputCode == "CC43 - Material" || item.CustomerOutputCode == "CC47 - Logistics" || item.CustomerOutputCode == "CC50 - Operation" || item.CustomerOutputCode == "CC52 - PM" || item.CustomerOutputCode == "CC53 - IT")
                    {
                        dic_totalPeopleAmount.Add(flag, item);
                        flags.Add(flag);
                    }
                    flag++;
                }

                FOL_Input_6_1 totalItem = new FOL_Input_6_1();
                FOL_Input_6_0_TotalMOH totalIDL = new FOL_Input_6_0_TotalMOH();

                List<FOL_Input_6_0_TotalMOH> list_TotalMOH = dc.FOL_Input_6_0_TotalMOH.Where(x => x.BPCOutputCode == "4606-IDL").ToList();//找到6.0这张表中IDL的Total值
                flag = 0;//用来标记每个core center的位置

                for (int i = 0; i < list1.Count; i++)
                {
                    if (list1[i].CustomerOutputCode == "CC31 - Facility" || list1[i].CustomerOutputCode == "CC33 - PE" || list1[i].CustomerOutputCode == "CC34 - Quality" || list1[i].CustomerOutputCode == "CC37 - Security"
                        || list1[i].CustomerOutputCode == "CC42 - Planning" || list1[i].CustomerOutputCode == "CC43 - Material" || list1[i].CustomerOutputCode == "CC47 - Logistics" || list1[i].CustomerOutputCode == "CC50 - Operation"
                        || list1[i].CustomerOutputCode == "CC52 - PM")
                    { continue; }
                    else if (list1[i].CustomerOutputCode == "CC53 - IT")
                    {
                        break;
                    }
                    else
                    {
                        if (i < flags[0]) { totalItem = dic_totalPeopleAmount[flags[0]]; totalIDL = list_TotalMOH[0]; }
                        else if (i > flags[0] && i < flags[1]) { totalItem = dic_totalPeopleAmount[flags[1]]; totalIDL = list_TotalMOH[1]; }
                        else if (i > flags[1] && i < flags[2]) { totalItem = dic_totalPeopleAmount[flags[2]]; totalIDL = list_TotalMOH[2]; }
                        else if (i > flags[2] && i < flags[3]) { totalItem = dic_totalPeopleAmount[flags[3]]; totalIDL = list_TotalMOH[3]; }
                        else if (i > flags[3] && i < flags[4]) { totalItem = dic_totalPeopleAmount[flags[4]]; totalIDL = list_TotalMOH[4]; }
                        else if (i > flags[4] && i < flags[5]) { totalItem = dic_totalPeopleAmount[flags[5]]; totalIDL = list_TotalMOH[5]; }
                        else if (i > flags[5] && i < flags[6]) { totalItem = dic_totalPeopleAmount[flags[6]]; totalIDL = list_TotalMOH[6]; }
                        else if (i > flags[6] && i < flags[7]) { totalItem = dic_totalPeopleAmount[flags[7]]; totalIDL = list_TotalMOH[7]; }
                        else if (i > flags[7] && i < flags[8]) { totalItem = dic_totalPeopleAmount[flags[8]]; totalIDL = list_TotalMOH[8]; }
                        else if (i > flags[8] && i < flags[9]) { totalItem = dic_totalPeopleAmount[flags[9]]; totalIDL = list_TotalMOH[9]; }
                        //else if (i > flags[9] && i < flags[10]) { totalItem = dic_totalPeopleAmount[flags[10]]; totalIDL = list_TotalMOH[10]; }
                        else { flag = i; break; };
                        FOL_Input_2_1 newItemPercent = new FOL_Input_2_1();
                        FOL_Input_1_1 newItemDollar = new FOL_Input_1_1();

                        //计算Percent
                        #region
                        newItemPercent.GLBPCCode = list1[i].GLBPCCode;
                        newItemPercent.GLBPCDescription = list1[i].GLBPCDescription;
                        newItemPercent.GLOutputCode = list1[i].GLOutputCode;
                        newItemPercent.CustomerBPCCode = list1[i].CustomerBPCCode;
                        newItemPercent.DIM1 = list1[i].DIM1;
                        newItemPercent.Order = list1[i].Order;
                        newItemPercent.CustomerOutputCode = list1[i].CustomerOutputCode;
                        newItemPercent.BU = list1[i].BU;
                        newItemPercent.Segment = list1[i].Segment;
                        newItemPercent.Period1 = (totalItem.Period1 == 0) ? 0 : (list1[i].Period1 / totalItem.Period1);
                        newItemPercent.Period2 = (totalItem.Period2 == 0) ? 0 : (list1[i].Period2 / totalItem.Period2);
                        newItemPercent.Period3 = (totalItem.Period3 == 0) ? 0 : (list1[i].Period3 / totalItem.Period3);
                        newItemPercent.Period4 = (totalItem.Period4 == 0) ? 0 : (list1[i].Period4 / totalItem.Period4);
                        newItemPercent.Period5 = (totalItem.Period5 == 0) ? 0 : (list1[i].Period5 / totalItem.Period5);
                        newItemPercent.Period6 = (totalItem.Period6 == 0) ? 0 : (list1[i].Period6 / totalItem.Period6);
                        newItemPercent.Period7 = (totalItem.Period7 == 0) ? 0 : (list1[i].Period7 / totalItem.Period7);
                        newItemPercent.Period8 = (totalItem.Period8 == 0) ? 0 : (list1[i].Period8 / totalItem.Period8);
                        newItemPercent.Period9 = (totalItem.Period9 == 0) ? 0 : (list1[i].Period9 / totalItem.Period9);
                        newItemPercent.Period10 = (totalItem.Period10 == 0) ? 0 : (list1[i].Period10 / totalItem.Period10);
                        newItemPercent.Period11 = (totalItem.Period11 == 0) ? 0 : (list1[i].Period11 / totalItem.Period11);
                        newItemPercent.Period12 = (totalItem.Period12 == 0) ? 0 : (list1[i].Period12 / totalItem.Period12);
                        newItemPercent.Period13 = (totalItem.Period13 == 0) ? 0 : (list1[i].Period13 / totalItem.Period13);
                        newItemPercent.Period14 = (totalItem.Period14 == 0) ? 0 : (list1[i].Period14 / totalItem.Period14);
                        newItemPercent.Period15 = (totalItem.Period15 == 0) ? 0 : (list1[i].Period15 / totalItem.Period15);
                        newItemPercent.Period16 = (totalItem.Period16 == 0) ? 0 : (list1[i].Period16 / totalItem.Period16);
                        newItemPercent.Type = type;
                        newItemPercent.InsertDate = System.DateTime.Now;
                        newItemPercent.InsertUser = "";
                        newItemPercent.Version = version;
                        #endregion
                        //计算Dollar Amount
                        #region
                        newItemDollar.GLBPCCode = list1[i].GLBPCCode;
                        newItemDollar.GLBPCDescription = list1[i].GLBPCDescription;
                        newItemDollar.GLOutputCode = list1[i].GLOutputCode;
                        newItemDollar.CustomerBPCCode = list1[i].CustomerBPCCode;
                        newItemDollar.DIM1 = list1[i].DIM1;
                        newItemDollar.Order = list1[i].Order;
                        newItemDollar.CustomerOutputCode = list1[i].CustomerOutputCode;
                        newItemDollar.BU = list1[i].BU;
                        newItemDollar.Segment = list1[i].Segment;
                        newItemDollar.Period1 = newItemPercent.Period1 * totalIDL.Period1;
                        newItemDollar.Period2 = newItemPercent.Period2 * totalIDL.Period2;
                        newItemDollar.Period3 = newItemPercent.Period3 * totalIDL.Period3;
                        newItemDollar.Period4 = newItemPercent.Period4 * totalIDL.Period4;
                        newItemDollar.Period5 = newItemPercent.Period5 * totalIDL.Period5;
                        newItemDollar.Period6 = newItemPercent.Period6 * totalIDL.Period6;
                        newItemDollar.Period7 = newItemPercent.Period7 * totalIDL.Period7;
                        newItemDollar.Period8 = newItemPercent.Period8 * totalIDL.Period8;
                        newItemDollar.Period9 = newItemPercent.Period9 * totalIDL.Period9;
                        newItemDollar.Period10 = newItemPercent.Period10 * totalIDL.Period10;
                        newItemDollar.Period11 = newItemPercent.Period11 * totalIDL.Period11;
                        newItemDollar.Period12 = newItemPercent.Period12 * totalIDL.Period12;
                        newItemDollar.Period13 = newItemPercent.Period13 * totalIDL.Period13;
                        newItemDollar.Period14 = newItemPercent.Period14 * totalIDL.Period14;
                        newItemDollar.Period15 = newItemPercent.Period15 * totalIDL.Period15;
                        newItemDollar.Period16 = newItemPercent.Period16 * totalIDL.Period16;
                        newItemDollar.Type = type;
                        newItemDollar.InsertDate = System.DateTime.Now;
                        newItemDollar.InsertUser = "";
                        newItemDollar.Version = version;
                        #endregion

                        list.Add(newItemPercent);
                        list_Dollar.Add(newItemDollar);
                    }
                }
                dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Version == version && x.Type == type));
                dc.FOL_Input_2_1.InsertAllOnSubmit(list);
                dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Version == version && x.Type == type));
                dc.FOL_Input_1_1.InsertAllOnSubmit(list_Dollar);
                dc.SubmitChanges();
                message = "Success";
                //Calculate SumPercent by project
                #region
                //int flag2 = 0;
                //for (int i = flag; i < list1.Count; i++)
                //{
                //    if (list1[i].GLBPCCode == "Total - MOH")
                //    {
                //        flag2 = i;
                //        break;
                //    }
                //    FOL_Input_2_1 projectTotal = new FOL_Input_2_1
                //    {
                //        BU = list1[i].BU,
                //        GLBPCCode = list1[i].GLBPCCode,
                //        CustomerBPCCode = list1[i].CustomerBPCCode,
                //        DIM1 = list1[i].DIM1,
                //        Order = list1[i].Order,
                //        Segment = list1[i].Segment,
                //        Period1 =0,
                //        Period2 = 0,
                //        Period3 = 0,
                //        Period4 = 0,
                //        Period5 = 0,
                //        Period6 = 0,
                //        Period7 = 0,
                //        Period8 = 0,
                //        Period9 = 0,
                //        Period10 = 0,
                //        Period11 = 0,
                //        Period12 = 0,
                //        Period13 = 0,
                //        Period14 = 0,
                //        Period15 = 0,
                //        Period16 = 0,
                //        Type = "6.1 IDL (%&#)",
                //        InsertDate = System.DateTime.Now,
                //        InsertUser = "",
                //        Version = version,
                //    };

                //    for(int j = 0; j < list.Count; j++)
                //    {
                //        if (list[j].GLBPCCode == list1[i].GLBPCCode)
                //        {
                //            projectTotal.Period1 = +list[j].Period1;
                //            projectTotal.Period2 = +list[j].Period2;
                //            projectTotal.Period3 = +list[j].Period3;
                //            projectTotal.Period4 = +list[j].Period4;
                //            projectTotal.Period5 = +list[j].Period5;
                //            projectTotal.Period6 = +list[j].Period6;
                //            projectTotal.Period7 = +list[j].Period7;
                //            projectTotal.Period8 = +list[j].Period8;
                //            projectTotal.Period9 = +list[j].Period9;
                //            projectTotal.Period10 = +list[j].Period10;
                //            projectTotal.Period11 = +list[j].Period11;
                //            projectTotal.Period12 = +list[j].Period12;
                //            projectTotal.Period13 = +list[j].Period13;
                //            projectTotal.Period14 = +list[j].Period14;
                //            projectTotal.Period15 = +list[j].Period15;
                //            projectTotal.Period16 = +list[j].Period16;
                //        }
                //    }
                //    list.Add(projectTotal);
                //}

                //for(int i = flag2; i < list1.Count; i++)
                //{
                //    FOL_Input_2_1 BUTotal = new FOL_Input_2_1
                //    {
                //        GLBPCCode = list1[i].GLBPCCode,
                //        GLBPCDescription = list1[i].GLBPCDescription,
                //        GLOutputCode = list1[i].GLOutputCode,
                //        CustomerBPCCode = list1[i].CustomerBPCCode,
                //        DIM1 = list1[i].DIM1,
                //        Order = list1[i].Order,
                //        CustomerOutputCode = list1[i].CustomerOutputCode,
                //        BU = list1[i].BU,
                //        Segment = list1[i].Segment,
                //        Period1 = 0,
                //        Period2 = 0,
                //        Period3 = 0,
                //        Period4 = 0,
                //        Period5 = 0,
                //        Period6 = 0,
                //        Period7 = 0,
                //        Period8 = 0,
                //        Period9 = 0,
                //        Period10 = 0,
                //        Period11 = 0,
                //        Period12 = 0,
                //        Period13 = 0,
                //        Period14 = 0,
                //        Period15 = 0,
                //        Period16 = 0,
                //        Type = "6.1 IDL (%&#)",
                //        InsertDate = System.DateTime.Now,
                //        InsertUser = "",
                //        Version = version,
                //    };
                //    for (int j = 0; j < list.Count; j++)
                //    {
                //        if (list[j].BU == list1[i].BU)
                //        {
                //            BUTotal.Period1 = +list[j].Period1;
                //            BUTotal.Period2 = +list[j].Period2;
                //            BUTotal.Period3 = +list[j].Period3;
                //            BUTotal.Period4 = +list[j].Period4;
                //            BUTotal.Period5 = +list[j].Period5;
                //            BUTotal.Period6 = +list[j].Period6;
                //            BUTotal.Period7 = +list[j].Period7;
                //            BUTotal.Period8 = +list[j].Period8;
                //            BUTotal.Period9 = +list[j].Period9;
                //            BUTotal.Period10 = +list[j].Period10;
                //            BUTotal.Period11 = +list[j].Period11;
                //            BUTotal.Period12 = +list[j].Period12;
                //            BUTotal.Period13 = +list[j].Period13;
                //            BUTotal.Period14 = +list[j].Period14;
                //            BUTotal.Period15 = +list[j].Period15;
                //            BUTotal.Period16 = +list[j].Period16;
                //        }
                //    }
                //    list.Add(BUTotal);
                //}
                #endregion
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return message;
        }

        //6.2, 6.3, 6.4, 6.5, 6.6
        public void Calculate6_2(string version, string type, List<FOL_Input_3_1> list_Amount, FOL_Input_3_1 total, FOL_Input_6_0_SumByType total_TypeAmount)
        {
            // Caculate percent
            List<FOL_Input_2_1> list = new List<FOL_Input_2_1>();
            List<FOL_Input_1_1> list_Dollar = new List<FOL_Input_1_1>();
            #region
            foreach (var item in list)
            {
                FOL_Input_2_1 newItem = new FOL_Input_2_1();
                FOL_Input_1_1 newItem_Dollar = new FOL_Input_1_1();

                //计算percent
                newItem.GLBPCCode = item.GLBPCCode;
                newItem.GLBPCDescription = item.GLBPCDescription;
                newItem.GLOutputCode = item.GLOutputCode;
                newItem.CustomerBPCCode = item.CustomerBPCCode;
                newItem.DIM1 = item.DIM1;
                newItem.Order = item.Order;
                newItem.CustomerOutputCode = item.CustomerOutputCode;
                newItem.BU = item.BU;
                newItem.Segment = item.Segment;
                newItem.Period1 = (total.Period1 == 0) ? 0 : (item.Period1 / total.Period1);
                newItem.Period2 = (total.Period1 == 0) ? 0 : (item.Period2 / total.Period2);
                newItem.Period3 = (total.Period1 == 0) ? 0 : (item.Period3 / total.Period3);
                newItem.Period4 = (total.Period1 == 0) ? 0 : (item.Period4 / total.Period4);
                newItem.Period5 = (total.Period1 == 0) ? 0 : (item.Period5 / total.Period5);
                newItem.Period6 = (total.Period1 == 0) ? 0 : (item.Period6 / total.Period6);
                newItem.Period7 = (total.Period1 == 0) ? 0 : (item.Period7 / total.Period7);
                newItem.Period8 = (total.Period1 == 0) ? 0 : (item.Period8 / total.Period8);
                newItem.Period9 = (total.Period1 == 0) ? 0 : (item.Period9 / total.Period9);
                newItem.Period10 = (total.Period1 == 0) ? 0 : (item.Period10 / total.Period10);
                newItem.Period11 = (total.Period1 == 0) ? 0 : (item.Period11 / total.Period11);
                newItem.Period12 = (total.Period1 == 0) ? 0 : (item.Period12 / total.Period12);
                newItem.Period13 = (total.Period1 == 0) ? 0 : (item.Period13 / total.Period13);
                newItem.Period14 = (total.Period1 == 0) ? 0 : (item.Period14 / total.Period14);
                newItem.Period15 = (total.Period1 == 0) ? 0 : (item.Period15 / total.Period15);
                newItem.Period16 = (total.Period1 == 0) ? 0 : (item.Period16 / total.Period16);
                newItem.Type = type;
                newItem.InsertDate = System.DateTime.Now;
                newItem.InsertUser = "";
                newItem.Version = version;
                list.Add(newItem);

                //计算dollar amount
                newItem_Dollar.GLBPCCode = item.GLBPCCode;
                newItem_Dollar.GLBPCDescription = item.GLBPCDescription;
                newItem_Dollar.GLOutputCode = item.GLOutputCode;
                newItem_Dollar.CustomerBPCCode = item.CustomerBPCCode;
                newItem_Dollar.DIM1 = item.DIM1;
                newItem_Dollar.Order = item.Order;
                newItem_Dollar.CustomerOutputCode = item.CustomerOutputCode;
                newItem_Dollar.BU = item.BU;
                newItem_Dollar.Segment = item.Segment;
                newItem_Dollar.Period1 = total_TypeAmount.Period1 * newItem.Period1;
                newItem_Dollar.Period2 = total_TypeAmount.Period2 * newItem.Period2;
                newItem_Dollar.Period3 = total_TypeAmount.Period3 * newItem.Period3;
                newItem_Dollar.Period4 = total_TypeAmount.Period4 * newItem.Period4;
                newItem_Dollar.Period5 = total_TypeAmount.Period5 * newItem.Period5;
                newItem_Dollar.Period6 = total_TypeAmount.Period6 * newItem.Period6;
                newItem_Dollar.Period7 = total_TypeAmount.Period7 * newItem.Period7;
                newItem_Dollar.Period8 = total_TypeAmount.Period8 * newItem.Period8;
                newItem_Dollar.Period9 = total_TypeAmount.Period9 * newItem.Period9;
                newItem_Dollar.Period10 = total_TypeAmount.Period10 * newItem.Period10;
                newItem_Dollar.Period11 = total_TypeAmount.Period11 * newItem.Period11;
                newItem_Dollar.Period12 = total_TypeAmount.Period12 * newItem.Period12;
                newItem_Dollar.Period13 = total_TypeAmount.Period13 * newItem.Period13;
                newItem_Dollar.Period14 = total_TypeAmount.Period14 * newItem.Period14;
                newItem_Dollar.Period15 = total_TypeAmount.Period15 * newItem.Period15;
                newItem_Dollar.Period16 = total_TypeAmount.Period16 * newItem.Period16;
                newItem_Dollar.Version = version;
                newItem_Dollar.Type = type;
                newItem_Dollar.InsertDate = System.DateTime.Now;
                newItem_Dollar.InsertUser = "";
                list_Dollar.Add(newItem_Dollar);
            }
            #endregion

            dc.FOL_Input_2_1.DeleteAllOnSubmit(dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList());
            dc.FOL_Input_2_1.InsertAllOnSubmit(list);
            dc.FOL_Input_1_1.DeleteAllOnSubmit(dc.FOL_Input_1_1.Where(x => x.Type == type && x.Version == version).ToList());
            dc.FOL_Input_1_1.InsertAllOnSubmit(list_Dollar);
            dc.SubmitChanges();
        }
        //7.0
        public List<FOL_Input_6_0_TotalMOH> Calculate7_0(string version)
        {
            List<FOL_Input_6_0_TotalMOH> list = new List<FOL_Input_6_0_TotalMOH>();
            List<string> list_department = dc.FOL_CCMapping.Where(x => x.Type != "MOH").Select(x => x.CCDescription).ToList();
            List<FOL_Input_6_0_SumByTypeModule> list_SumByTypeModule = dc.FOL_Input_6_0_SumByTypeModule.ToList();
            List<FOL_Input_6_0_SumByType> list_SumByType = new List<FOL_Input_6_0_SumByType>();
            for (int n = 0; n < list_SumByTypeModule.Count; n++)
            {
                string InsertType = "";
                if (n > 5) InsertType = "7.2 S.M.(%)";
                else InsertType = "7.1 G.A.(%)";
                FOL_Input_6_0_SumByType newItem_SumByType = new FOL_Input_6_0_SumByType();
                newItem_SumByType.GLBPCCode = list_SumByTypeModule[n].GLBPCCode;
                newItem_SumByType.GLBPCDescription = list_SumByTypeModule[n].GLBPCDescription;
                newItem_SumByType.GLOutputCode = list_SumByTypeModule[n].GLOutputCode;
                newItem_SumByType.PARENTH1 = list_SumByTypeModule[n].PARENTH1;
                newItem_SumByType.UploadCode = list_SumByTypeModule[n].UploadCode;
                newItem_SumByType.BPCOutputCode = list_SumByTypeModule[n].BPCOutputCode;
                newItem_SumByType.BU = list_SumByTypeModule[n].BU;
                newItem_SumByType.Segment = list_SumByTypeModule[n].Segment;
                newItem_SumByType.Period1 = 0;
                newItem_SumByType.Period2 = 0;
                newItem_SumByType.Period3 = 0;
                newItem_SumByType.Period4 = 0;
                newItem_SumByType.Period5 = 0;
                newItem_SumByType.Period6 = 0;
                newItem_SumByType.Period7 = 0;
                newItem_SumByType.Period8 = 0;
                newItem_SumByType.Period9 = 0;
                newItem_SumByType.Period10 = 0;
                newItem_SumByType.Period11 = 0;
                newItem_SumByType.Period12 = 0;
                newItem_SumByType.Period13 = 0;
                newItem_SumByType.Period14 = 0;
                newItem_SumByType.Period15 = 0;
                newItem_SumByType.Period16 = 0;
                newItem_SumByType.Type = InsertType;
                newItem_SumByType.Version = version;
                newItem_SumByType.InsertDate = System.DateTime.Now;
                newItem_SumByType.InsertUser = "";
                list_SumByType.Add(newItem_SumByType);
            }
            for (int n = 6; n < list_SumByTypeModule.Count * 2; n++)
            {
                string InsertType = "";
                if (n > 5) InsertType = "7.2 S.M.(%)";
                else InsertType = "7.1 G.A.(%)";
                FOL_Input_6_0_SumByType newItem_SumByType = new FOL_Input_6_0_SumByType();
                newItem_SumByType.GLBPCCode = list_SumByTypeModule[n - 6].GLBPCCode;
                newItem_SumByType.GLBPCDescription = list_SumByTypeModule[n - 6].GLBPCDescription;
                newItem_SumByType.GLOutputCode = list_SumByTypeModule[n - 6].GLOutputCode;
                newItem_SumByType.PARENTH1 = list_SumByTypeModule[n - 6].PARENTH1;
                newItem_SumByType.UploadCode = list_SumByTypeModule[n - 6].UploadCode;
                newItem_SumByType.BPCOutputCode = list_SumByTypeModule[n - 6].BPCOutputCode;
                newItem_SumByType.BU = list_SumByTypeModule[n - 6].BU;
                newItem_SumByType.Segment = list_SumByTypeModule[n - 6].Segment;
                newItem_SumByType.Period1 = 0;
                newItem_SumByType.Period2 = 0;
                newItem_SumByType.Period3 = 0;
                newItem_SumByType.Period4 = 0;
                newItem_SumByType.Period5 = 0;
                newItem_SumByType.Period6 = 0;
                newItem_SumByType.Period7 = 0;
                newItem_SumByType.Period8 = 0;
                newItem_SumByType.Period9 = 0;
                newItem_SumByType.Period10 = 0;
                newItem_SumByType.Period11 = 0;
                newItem_SumByType.Period12 = 0;
                newItem_SumByType.Period13 = 0;
                newItem_SumByType.Period14 = 0;
                newItem_SumByType.Period15 = 0;
                newItem_SumByType.Period16 = 0;
                newItem_SumByType.Type = InsertType;
                newItem_SumByType.Version = version;
                newItem_SumByType.InsertDate = System.DateTime.Now;
                newItem_SumByType.InsertUser = "";
                list_SumByType.Add(newItem_SumByType);
            }
            for (int i = 0; i < list_department.Count; i++)
            {
                string department = list_department[i];
                List<FOL_Input_6_0_TotalMOH> list_new = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department == department && x.IsSubtotal != 11).ToList();
                List<FOL_Input_6_0_TotalMOH> list_byDepAndBPCCode = dc.FOL_Input_6_0_TotalMOH.Where(x => x.Department == department && x.IsSubtotal == 11).ToList();

                FOL_Input_6_0_TotalMOH newItem = new FOL_Input_6_0_TotalMOH();

                for (int j = 0; j < list_new.Count; j++)
                {
                    //计算总和By BPCCode By Dep
                    #region
                    if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[0].PARENTH1)
                    {
                        list_byDepAndBPCCode[0].Period1 = (list_byDepAndBPCCode[0].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[0].Period2 = (list_byDepAndBPCCode[0].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[0].Period3 = (list_byDepAndBPCCode[0].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[0].Period4 = (list_byDepAndBPCCode[0].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[0].Period5 = (list_byDepAndBPCCode[0].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[0].Period6 = (list_byDepAndBPCCode[0].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[0].Period7 = (list_byDepAndBPCCode[0].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[0].Period8 = (list_byDepAndBPCCode[0].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[0].Period9 = (list_byDepAndBPCCode[0].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[0].Period10 = (list_byDepAndBPCCode[0].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[0].Period11 = (list_byDepAndBPCCode[0].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[0].Period12 = (list_byDepAndBPCCode[0].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[0].Period13 = (list_byDepAndBPCCode[0].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[0].Period14 = (list_byDepAndBPCCode[0].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[0].Period15 = (list_byDepAndBPCCode[0].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[0].Period16 = (list_byDepAndBPCCode[0].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[0].Period1 = (list_SumByType[0].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[0].Period2 = (list_SumByType[0].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[0].Period3 = (list_SumByType[0].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[0].Period4 = (list_SumByType[0].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[0].Period5 = (list_SumByType[0].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[0].Period6 = (list_SumByType[0].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[0].Period7 = (list_SumByType[0].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[0].Period8 = (list_SumByType[0].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[0].Period9 = (list_SumByType[0].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[0].Period10 = (list_SumByType[0].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[0].Period11 = (list_SumByType[0].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[0].Period12 = (list_SumByType[0].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[0].Period13 = (list_SumByType[0].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[0].Period14 = (list_SumByType[0].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[0].Period15 = (list_SumByType[0].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[0].Period16 = (list_SumByType[0].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[6].Period1 = (list_SumByType[6].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[6].Period2 = (list_SumByType[6].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[6].Period3 = (list_SumByType[6].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[6].Period4 = (list_SumByType[6].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[6].Period5 = (list_SumByType[6].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[6].Period6 = (list_SumByType[6].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[6].Period7 = (list_SumByType[6].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[6].Period8 = (list_SumByType[6].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[6].Period9 = (list_SumByType[6].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[6].Period10 = (list_SumByType[6].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[6].Period11 = (list_SumByType[6].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[6].Period12 = (list_SumByType[6].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[6].Period13 = (list_SumByType[6].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[6].Period14 = (list_SumByType[6].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[6].Period15 = (list_SumByType[6].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[6].Period16 = (list_SumByType[6].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[1].PARENTH1)
                    {
                        list_byDepAndBPCCode[1].Period1 = (list_byDepAndBPCCode[1].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[1].Period2 = (list_byDepAndBPCCode[1].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[1].Period3 = (list_byDepAndBPCCode[1].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[1].Period4 = (list_byDepAndBPCCode[1].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[1].Period5 = (list_byDepAndBPCCode[1].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[1].Period6 = (list_byDepAndBPCCode[1].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[1].Period7 = (list_byDepAndBPCCode[1].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[1].Period8 = (list_byDepAndBPCCode[1].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[1].Period9 = (list_byDepAndBPCCode[1].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[1].Period10 = (list_byDepAndBPCCode[1].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[1].Period11 = (list_byDepAndBPCCode[1].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[1].Period12 = (list_byDepAndBPCCode[1].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[1].Period13 = (list_byDepAndBPCCode[1].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[1].Period14 = (list_byDepAndBPCCode[1].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[1].Period15 = (list_byDepAndBPCCode[1].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[1].Period16 = (list_byDepAndBPCCode[1].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[1].Period1 = (list_SumByType[1].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[1].Period2 = (list_SumByType[1].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[1].Period3 = (list_SumByType[1].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[1].Period4 = (list_SumByType[1].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[1].Period5 = (list_SumByType[1].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[1].Period6 = (list_SumByType[1].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[1].Period7 = (list_SumByType[1].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[1].Period8 = (list_SumByType[1].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[1].Period9 = (list_SumByType[1].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[1].Period10 = (list_SumByType[1].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[1].Period11 = (list_SumByType[1].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[1].Period12 = (list_SumByType[1].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[1].Period13 = (list_SumByType[1].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[1].Period14 = (list_SumByType[1].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[1].Period15 = (list_SumByType[1].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[1].Period16 = (list_SumByType[1].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[7].Period1 = (list_SumByType[7].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[7].Period2 = (list_SumByType[7].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[7].Period3 = (list_SumByType[7].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[7].Period4 = (list_SumByType[7].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[7].Period5 = (list_SumByType[7].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[7].Period6 = (list_SumByType[7].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[7].Period7 = (list_SumByType[7].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[7].Period8 = (list_SumByType[7].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[7].Period9 = (list_SumByType[7].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[7].Period10 = (list_SumByType[7].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[7].Period11 = (list_SumByType[7].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[7].Period12 = (list_SumByType[7].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[7].Period13 = (list_SumByType[7].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[7].Period14 = (list_SumByType[7].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[7].Period15 = (list_SumByType[7].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[7].Period16 = (list_SumByType[7].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[2].PARENTH1)
                    {
                        list_byDepAndBPCCode[2].Period1 = (list_byDepAndBPCCode[2].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[2].Period2 = (list_byDepAndBPCCode[2].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[2].Period3 = (list_byDepAndBPCCode[2].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[2].Period4 = (list_byDepAndBPCCode[2].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[2].Period5 = (list_byDepAndBPCCode[2].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[2].Period6 = (list_byDepAndBPCCode[2].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[2].Period7 = (list_byDepAndBPCCode[2].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[2].Period8 = (list_byDepAndBPCCode[2].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[2].Period9 = (list_byDepAndBPCCode[2].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[2].Period10 = (list_byDepAndBPCCode[2].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[2].Period11 = (list_byDepAndBPCCode[2].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[2].Period12 = (list_byDepAndBPCCode[2].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[2].Period13 = (list_byDepAndBPCCode[2].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[2].Period14 = (list_byDepAndBPCCode[2].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[2].Period15 = (list_byDepAndBPCCode[2].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[2].Period16 = (list_byDepAndBPCCode[2].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[2].Period1 = (list_SumByType[2].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[2].Period2 = (list_SumByType[2].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[2].Period3 = (list_SumByType[2].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[2].Period4 = (list_SumByType[2].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[2].Period5 = (list_SumByType[2].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[2].Period6 = (list_SumByType[2].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[2].Period7 = (list_SumByType[2].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[2].Period8 = (list_SumByType[2].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[2].Period9 = (list_SumByType[2].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[2].Period10 = (list_SumByType[2].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[2].Period11 = (list_SumByType[2].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[2].Period12 = (list_SumByType[2].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[2].Period13 = (list_SumByType[2].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[2].Period14 = (list_SumByType[2].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[2].Period15 = (list_SumByType[2].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[2].Period16 = (list_SumByType[2].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[8].Period1 = (list_SumByType[8].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[8].Period2 = (list_SumByType[8].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[8].Period3 = (list_SumByType[8].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[8].Period4 = (list_SumByType[8].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[8].Period5 = (list_SumByType[8].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[8].Period6 = (list_SumByType[8].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[8].Period7 = (list_SumByType[8].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[8].Period8 = (list_SumByType[8].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[8].Period9 = (list_SumByType[8].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[8].Period10 = (list_SumByType[8].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[8].Period11 = (list_SumByType[8].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[8].Period12 = (list_SumByType[8].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[8].Period13 = (list_SumByType[8].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[8].Period14 = (list_SumByType[8].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[8].Period15 = (list_SumByType[8].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[8].Period16 = (list_SumByType[8].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[3].PARENTH1)
                    {
                        list_byDepAndBPCCode[3].Period1 = (list_byDepAndBPCCode[3].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[3].Period2 = (list_byDepAndBPCCode[3].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[3].Period3 = (list_byDepAndBPCCode[3].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[3].Period4 = (list_byDepAndBPCCode[3].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[3].Period5 = (list_byDepAndBPCCode[3].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[3].Period6 = (list_byDepAndBPCCode[3].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[3].Period7 = (list_byDepAndBPCCode[3].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[3].Period8 = (list_byDepAndBPCCode[3].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[3].Period9 = (list_byDepAndBPCCode[3].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[3].Period10 = (list_byDepAndBPCCode[3].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[3].Period11 = (list_byDepAndBPCCode[3].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[3].Period12 = (list_byDepAndBPCCode[3].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[3].Period13 = (list_byDepAndBPCCode[3].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[3].Period14 = (list_byDepAndBPCCode[3].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[3].Period15 = (list_byDepAndBPCCode[3].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[3].Period16 = (list_byDepAndBPCCode[3].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[3].Period1 = (list_SumByType[3].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[3].Period2 = (list_SumByType[3].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[3].Period3 = (list_SumByType[3].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[3].Period4 = (list_SumByType[3].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[3].Period5 = (list_SumByType[3].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[3].Period6 = (list_SumByType[3].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[3].Period7 = (list_SumByType[3].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[3].Period8 = (list_SumByType[3].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[3].Period9 = (list_SumByType[3].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[3].Period10 = (list_SumByType[3].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[3].Period11 = (list_SumByType[3].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[3].Period12 = (list_SumByType[3].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[3].Period13 = (list_SumByType[3].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[3].Period14 = (list_SumByType[3].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[3].Period15 = (list_SumByType[3].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[3].Period16 = (list_SumByType[3].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[9].Period1 = (list_SumByType[9].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[9].Period2 = (list_SumByType[9].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[9].Period3 = (list_SumByType[9].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[9].Period4 = (list_SumByType[9].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[9].Period5 = (list_SumByType[9].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[9].Period6 = (list_SumByType[9].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[9].Period7 = (list_SumByType[9].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[9].Period8 = (list_SumByType[9].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[9].Period9 = (list_SumByType[9].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[9].Period10 = (list_SumByType[9].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[9].Period11 = (list_SumByType[9].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[9].Period12 = (list_SumByType[9].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[9].Period13 = (list_SumByType[9].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[9].Period14 = (list_SumByType[9].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[9].Period15 = (list_SumByType[9].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[9].Period16 = (list_SumByType[9].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[4].PARENTH1)
                    {
                        list_byDepAndBPCCode[4].Period1 = (list_byDepAndBPCCode[4].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[4].Period2 = (list_byDepAndBPCCode[4].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[4].Period3 = (list_byDepAndBPCCode[4].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[4].Period4 = (list_byDepAndBPCCode[4].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[4].Period5 = (list_byDepAndBPCCode[4].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[4].Period6 = (list_byDepAndBPCCode[4].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[4].Period7 = (list_byDepAndBPCCode[4].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[4].Period8 = (list_byDepAndBPCCode[4].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[4].Period9 = (list_byDepAndBPCCode[4].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[4].Period10 = (list_byDepAndBPCCode[4].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[4].Period11 = (list_byDepAndBPCCode[4].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[4].Period12 = (list_byDepAndBPCCode[4].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[4].Period13 = (list_byDepAndBPCCode[4].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[4].Period14 = (list_byDepAndBPCCode[4].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[4].Period15 = (list_byDepAndBPCCode[4].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[4].Period16 = (list_byDepAndBPCCode[4].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[4].Period1 = (list_SumByType[4].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[4].Period2 = (list_SumByType[4].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[4].Period3 = (list_SumByType[4].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[4].Period4 = (list_SumByType[4].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[4].Period5 = (list_SumByType[4].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[4].Period6 = (list_SumByType[4].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[4].Period7 = (list_SumByType[4].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[4].Period8 = (list_SumByType[4].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[4].Period9 = (list_SumByType[4].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[4].Period10 = (list_SumByType[4].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[4].Period11 = (list_SumByType[4].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[4].Period12 = (list_SumByType[4].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[4].Period13 = (list_SumByType[4].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[4].Period14 = (list_SumByType[4].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[4].Period15 = (list_SumByType[4].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[4].Period16 = (list_SumByType[1].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[10].Period1 = (list_SumByType[10].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[10].Period2 = (list_SumByType[10].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[10].Period3 = (list_SumByType[10].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[10].Period4 = (list_SumByType[10].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[10].Period5 = (list_SumByType[10].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[10].Period6 = (list_SumByType[10].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[10].Period7 = (list_SumByType[10].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[10].Period8 = (list_SumByType[10].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[10].Period9 = (list_SumByType[10].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[10].Period10 = (list_SumByType[10].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[10].Period11 = (list_SumByType[10].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[10].Period12 = (list_SumByType[10].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[10].Period13 = (list_SumByType[10].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[10].Period14 = (list_SumByType[10].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[10].Period15 = (list_SumByType[10].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[10].Period16 = (list_SumByType[10].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    else if (list_new[j].BPCOutputCode == list_byDepAndBPCCode[5].PARENTH1)
                    {
                        list_byDepAndBPCCode[5].Period1 = (list_byDepAndBPCCode[5].Period1 ?? 0) + list_new[j].Period1;
                        list_byDepAndBPCCode[5].Period2 = (list_byDepAndBPCCode[5].Period2 ?? 0) + list_new[j].Period2;
                        list_byDepAndBPCCode[5].Period3 = (list_byDepAndBPCCode[5].Period3 ?? 0) + list_new[j].Period3;
                        list_byDepAndBPCCode[5].Period4 = (list_byDepAndBPCCode[5].Period4 ?? 0) + list_new[j].Period4;
                        list_byDepAndBPCCode[5].Period5 = (list_byDepAndBPCCode[5].Period5 ?? 0) + list_new[j].Period5;
                        list_byDepAndBPCCode[5].Period6 = (list_byDepAndBPCCode[5].Period6 ?? 0) + list_new[j].Period6;
                        list_byDepAndBPCCode[5].Period7 = (list_byDepAndBPCCode[5].Period7 ?? 0) + list_new[j].Period7;
                        list_byDepAndBPCCode[5].Period8 = (list_byDepAndBPCCode[5].Period8 ?? 0) + list_new[j].Period8;
                        list_byDepAndBPCCode[5].Period9 = (list_byDepAndBPCCode[5].Period9 ?? 0) + list_new[j].Period9;
                        list_byDepAndBPCCode[5].Period10 = (list_byDepAndBPCCode[5].Period10 ?? 0) + list_new[j].Period10;
                        list_byDepAndBPCCode[5].Period11 = (list_byDepAndBPCCode[5].Period11 ?? 0) + list_new[j].Period11;
                        list_byDepAndBPCCode[5].Period12 = (list_byDepAndBPCCode[5].Period12 ?? 0) + list_new[j].Period12;
                        list_byDepAndBPCCode[5].Period13 = (list_byDepAndBPCCode[5].Period13 ?? 0) + list_new[j].Period13;
                        list_byDepAndBPCCode[5].Period14 = (list_byDepAndBPCCode[5].Period14 ?? 0) + list_new[j].Period14;
                        list_byDepAndBPCCode[5].Period15 = (list_byDepAndBPCCode[5].Period15 ?? 0) + list_new[j].Period15;
                        list_byDepAndBPCCode[5].Period16 = (list_byDepAndBPCCode[5].Period16 ?? 0) + list_new[j].Period16;

                        if (list_new[j].Department != "Business Development")
                        {
                            list_SumByType[5].Period1 = (list_SumByType[5].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[5].Period2 = (list_SumByType[5].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[5].Period3 = (list_SumByType[5].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[5].Period4 = (list_SumByType[5].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[5].Period5 = (list_SumByType[5].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[5].Period6 = (list_SumByType[5].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[5].Period7 = (list_SumByType[5].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[5].Period8 = (list_SumByType[5].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[5].Period9 = (list_SumByType[5].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[5].Period10 = (list_SumByType[5].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[5].Period11 = (list_SumByType[5].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[5].Period12 = (list_SumByType[5].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[5].Period13 = (list_SumByType[5].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[5].Period14 = (list_SumByType[5].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[5].Period15 = (list_SumByType[5].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[5].Period16 = (list_SumByType[5].Period16 ?? 0) + list_new[j].Period16;
                        }
                        else
                        {
                            list_SumByType[11].Period1 = (list_SumByType[11].Period1 ?? 0) + list_new[j].Period1;
                            list_SumByType[11].Period2 = (list_SumByType[11].Period2 ?? 0) + list_new[j].Period2;
                            list_SumByType[11].Period3 = (list_SumByType[11].Period3 ?? 0) + list_new[j].Period3;
                            list_SumByType[11].Period4 = (list_SumByType[11].Period4 ?? 0) + list_new[j].Period4;
                            list_SumByType[11].Period5 = (list_SumByType[11].Period5 ?? 0) + list_new[j].Period5;
                            list_SumByType[11].Period6 = (list_SumByType[11].Period6 ?? 0) + list_new[j].Period6;
                            list_SumByType[11].Period7 = (list_SumByType[11].Period7 ?? 0) + list_new[j].Period7;
                            list_SumByType[11].Period8 = (list_SumByType[11].Period8 ?? 0) + list_new[j].Period8;
                            list_SumByType[11].Period9 = (list_SumByType[11].Period9 ?? 0) + list_new[j].Period9;
                            list_SumByType[11].Period10 = (list_SumByType[11].Period10 ?? 0) + list_new[j].Period10;
                            list_SumByType[11].Period11 = (list_SumByType[11].Period11 ?? 0) + list_new[j].Period11;
                            list_SumByType[11].Period12 = (list_SumByType[11].Period12 ?? 0) + list_new[j].Period12;
                            list_SumByType[11].Period13 = (list_SumByType[11].Period13 ?? 0) + list_new[j].Period13;
                            list_SumByType[11].Period14 = (list_SumByType[11].Period14 ?? 0) + list_new[j].Period14;
                            list_SumByType[11].Period15 = (list_SumByType[11].Period15 ?? 0) + list_new[j].Period15;
                            list_SumByType[11].Period16 = (list_SumByType[11].Period16 ?? 0) + list_new[j].Period16;
                        }
                    }
                    #endregion
                    //计算IsSubtotal=1的类型的和
                    #region
                    if (list_new[j].IsSubtotal == 1)
                    {
                        list_new[j].Period1 = newItem.Period1;
                        list_new[j].Period2 = newItem.Period2;
                        list_new[j].Period3 = newItem.Period3;
                        list_new[j].Period4 = newItem.Period4;
                        list_new[j].Period5 = newItem.Period5;
                        list_new[j].Period6 = newItem.Period6;
                        list_new[j].Period7 = newItem.Period7;
                        list_new[j].Period8 = newItem.Period8;
                        list_new[j].Period9 = newItem.Period9;
                        list_new[j].Period10 = newItem.Period10;
                        list_new[j].Period11 = newItem.Period11;
                        list_new[j].Period12 = newItem.Period12;
                        list_new[j].Period13 = newItem.Period13;
                        list_new[j].Period14 = newItem.Period14;
                        list_new[j].Period15 = newItem.Period15;
                        list_new[j].Period16 = newItem.Period16;
                        newItem = new FOL_Input_6_0_TotalMOH();
                    }
                    else
                    {
                        newItem.Period1 = (newItem.Period1 ?? 0) + list_new[j].Period1;
                        newItem.Period2 = (newItem.Period2 ?? 0) + list_new[j].Period2;
                        newItem.Period3 = (newItem.Period3 ?? 0) + list_new[j].Period3;
                        newItem.Period4 = (newItem.Period4 ?? 0) + list_new[j].Period4;
                        newItem.Period5 = (newItem.Period5 ?? 0) + list_new[j].Period5;
                        newItem.Period6 = (newItem.Period6 ?? 0) + list_new[j].Period6;
                        newItem.Period7 = (newItem.Period7 ?? 0) + list_new[j].Period7;
                        newItem.Period8 = (newItem.Period8 ?? 0) + list_new[j].Period8;
                        newItem.Period9 = (newItem.Period9 ?? 0) + list_new[j].Period9;
                        newItem.Period10 = (newItem.Period10 ?? 0) + list_new[j].Period10;
                        newItem.Period11 = (newItem.Period11 ?? 0) + list_new[j].Period11;
                        newItem.Period12 = (newItem.Period12 ?? 0) + list_new[j].Period12;
                        newItem.Period13 = (newItem.Period13 ?? 0) + list_new[j].Period13;
                        newItem.Period14 = (newItem.Period14 ?? 0) + list_new[j].Period14;
                        newItem.Period15 = (newItem.Period15 ?? 0) + list_new[j].Period15;
                        newItem.Period16 = (newItem.Period16 ?? 0) + list_new[j].Period16;
                    }
                    #endregion

                }
            }
            dc.FOL_Input_6_0_SumByType.DeleteAllOnSubmit(dc.FOL_Input_6_0_SumByType.Where(x => (x.Type == "7.1 G.A.(%)" || x.Type == "7.2 S.M.(%)") && x.Version == version).ToList());
            dc.FOL_Input_6_0_SumByType.InsertAllOnSubmit(list_SumByType);

            dc.SubmitChanges();

            return list;
        }

        //计算7.1
        public List<FOL_Input_1_1> Calculate7_1(string version, string type)
        {
            List<FOL_Input_1_1> list = dc.FOL_Input_1_1.Where(x => x.Type == type && x.Version == version).ToList();
            List<FOL_Input_6_0_SumByType> list_SumByType = dc.FOL_Input_6_0_SumByType.Where(x => x.Version == version && x.Type == type).ToList();

            int j = 0;
            List<FOL_Input_2_1> list_percent = dc.FOL_Input_2_1.Where(x => x.Type == type && x.Version == version).ToList();
            int flag = 0;
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].CustomerOutputCode.ToUpper().Contains("CCGA"))
                {
                    j = 0;
                    flag++;
                    continue;
                }
                list[i].Period1 = list_SumByType[flag].Period1 * (list_percent[j].Period1 ?? 0);
                list[i].Period2 = list_SumByType[flag].Period2 * (list_percent[j].Period2 ?? 0);
                list[i].Period3 = list_SumByType[flag].Period3 * (list_percent[j].Period3 ?? 0);
                list[i].Period4 = list_SumByType[flag].Period4 * (list_percent[j].Period4 ?? 0);
                list[i].Period5 = list_SumByType[flag].Period5 * (list_percent[j].Period5 ?? 0);
                list[i].Period6 = list_SumByType[flag].Period6 * (list_percent[j].Period6 ?? 0);
                list[i].Period7 = list_SumByType[flag].Period7 * (list_percent[j].Period7 ?? 0);
                list[i].Period8 = list_SumByType[flag].Period8 * (list_percent[j].Period8 ?? 0);
                list[i].Period9 = list_SumByType[flag].Period9 * (list_percent[j].Period9 ?? 0);
                list[i].Period10 = list_SumByType[flag].Period10 * (list_percent[j].Period10 ?? 0);
                list[i].Period11 = list_SumByType[flag].Period11 * (list_percent[j].Period11 ?? 0);
                list[i].Period12 = list_SumByType[flag].Period12 * (list_percent[j].Period12 ?? 0);
                list[i].Period13 = list_SumByType[flag].Period13 * (list_percent[j].Period13 ?? 0);
                list[i].Period14 = list_SumByType[flag].Period14 * (list_percent[j].Period14 ?? 0);
                list[i].Period15 = list_SumByType[flag].Period15 * (list_percent[j].Period15 ?? 0);
                list[i].Period16 = list_SumByType[flag].Period16 * (list_percent[j].Period16 ?? 0);
                j++;
            }
            dc.SubmitChanges();
            return list;
        }

        //计算7.1, 7.2的percent
        public List<FOL_Input_2_1> Calculate7_1_Percent(string version, string type)
        {
            //Caculate percent
            List<FOL_Input_1_1> projectList1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> projectList2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> projectList3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_2_1> percentList2_1 = dc.FOL_Input_2_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            List<FOL_Input_1_1> projectList2_1 = dc.FOL_Input_1_1.Where(x => x.Type == "2.1 Std VAM % from Ops(%)" && x.Version == version).ToList();
            List<FOL_Input_1_1> projectList1_0 = dc.FOL_Input_1_1.Where(x => x.Type == "1.0 Total Sales($)" && x.Version == version).ToList();

            FOL_Input_1_1 site1 = new FOL_Input_1_1();
            FOL_Input_1_1 site2 = new FOL_Input_1_1();
            FOL_Input_1_1 site3 = new FOL_Input_1_1();
            FOL_Input_1_1 site1_0 = new FOL_Input_1_1();
            FOL_Input_1_1 site2_1 = new FOL_Input_1_1();
            FOL_Input_2_1 sitePercent2_1 = new FOL_Input_2_1();

            //*********site1,2,3的一些赋值
            #region
            site1.GLBPCCode = "3000";
            site1.GLBPCDescription = "Gross Sales: External";
            site1.GLOutputCode = "3000 - Site";
            site1.CustomerOutputCode = "Total - Site";
            site1.Segment = "Site";
            site1.Version = version;
            site1.InsertDate = System.DateTime.Now;

            site2.GLBPCCode = "3500";
            site2.GLBPCDescription = "Gross Sales: Intercompany";
            site2.GLOutputCode = "3500 - Site";
            site2.CustomerOutputCode = "Total - Site";
            site2.Segment = "Site";
            site2.Version = version;
            site2.InsertDate = System.DateTime.Now;

            site3.GLBPCCode = "3010";
            site3.GLBPCDescription = "Non-Recurring Sales Activities and Other Cost Recovery";
            site3.GLOutputCode = "3010 - Site";
            site3.CustomerOutputCode = "Total - Site";
            site3.Segment = "Site";
            site3.Version = version;
            site3.InsertDate = System.DateTime.Now;

            site1_0.GLBPCCode = "30_SALES";
            site1_0.GLBPCDescription = "Total Sales";
            site1_0.GLOutputCode = "30_SALES - AGI";
            site1_0.CustomerOutputCode = "Total - Site";
            site1_0.Segment = "Site";
            site1_0.Version = version;
            site1_0.InsertDate = System.DateTime.Now;
            #endregion

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
            #endregion
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

            //********计算percent= {Project (1.1 + 1.2 + 1.3) ÷ Site (1.1 + 1.2 + 1.3)+ [Project(1.1 + 1.2) @ Project(2.1 %) + Project(1.3)] ÷ [Site(1.1 + 1.2) @ Site(2.1 %) + Site(1.3)]} ÷ 2
            for (int i = 0; i < projectList1.Count; i++)
            {
                //FOL_Input_1_1 newItem_ProjectSumOfThree = new FOL_Input_1_1();
                //FOL_Input_1_1 newItem_ProjectSumOfTwo = new FOL_Input_1_1();
                //FOL_Input_1_1 newItem_SiteSumOfTwo = new FOL_Input_1_1();
                FOL_Input_2_1 newItem_Result = new FOL_Input_2_1();
                newItem_Result.CustomerBPCCode = projectList1[i].CustomerBPCCode;
                newItem_Result.CustomerOutputCode = projectList1[i].CustomerOutputCode;
                newItem_Result.DIM1 = projectList1[i].DIM1;
                newItem_Result.Order = projectList1[i].Order;
                newItem_Result.BU = projectList1[i].BU;
                newItem_Result.Segment = projectList1[i].Segment;
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
                newItem_Result.Type = type;
                newItem_Result.Version = version;
                newItem_Result.InsertUser = "";
                newItem_Result.InsertDate = System.DateTime.Now;
                listResult.Add(newItem_Result);
            }

            return listResult;
        }
        //8.0
        public string Calculate8_0(string version,string type)
        {
            string message = "";
            List<FOL_Input_1_1> list1 = dc.FOL_Input_1_1.Where(x => x.Type == "1.1 Gross Sales - External($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list2 = dc.FOL_Input_1_1.Where(x => x.Type == "1.2 Gross Sales - Interco($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list3 = dc.FOL_Input_1_1.Where(x => x.Type == "1.3 Gross Sales - Recoveries($)" && x.Version == version).ToList();
            List<FOL_Input_1_1> list4 = dc.FOL_Input_1_1.Where(x => x.Type == "1.4 Rev Recog - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_1_1> list5 = dc.FOL_Input_1_1.Where(x => x.Type == "1.5 Rev Reversal - OT Contract" && x.Version == version).ToList();
            List<FOL_Input_2_1> list6 = dc.FOL_Input_2_1.Where(x => x.Type == "8.0 Corp. Alloc. (%)" && x.Version == version).ToList();
            if (list1.Count==0)
            { message = "该月份的1.1 Gross Sales - External($)数据为空;"; return message; }
            if (list2.Count==0)
            { message = message + "该月份的1.2 Gross Sales - Interco($)数据为空;"; return message; }
            if (list3.Count==0)
            { message = message + "该月份的1.3 Gross Sales - Recoveries($)数据为空;"; return message; }
            if (list4.Count==0)
            { message = message + "该月份的1.4 Rev Recog - OT Contract数据为空;"; return message; }
            if (list5.Count==0)
            { message = message + "该月份的1.5 Rev Reversal - OT Contract数据为空;"; return message; }
            if (list6.Count==0)
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