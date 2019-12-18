using ApiForRackManage.DataContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

namespace ApiForRackManage.Controllers
{
    public class SystemController : ApiController
    {
        RackDataContext dc = new RackDataContext();
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);
        public IDbConnection connPNINOUT = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);
        public IDbConnection connBaan = new SqlConnection(ConfigurationManager.ConnectionStrings["Baan"].ConnectionString);

        [HttpGet]
        public Resp getSystemList()
        {
            Resp resp = new Resp();
            List<Instrument> listUF = new List<Instrument>();
            List<J750_Board> list750 = new List<J750_Board>();
            listUF = dc.Instrument.ToList();
            list750 = dc.J750_Board.ToList();
            resp.Code = 200;
            resp.DataUF = listUF;
            resp.Data750 = list750;
            resp.Message = "success";
            return resp;
        }

        [HttpPost]
        public int UpdateSystem(SysList list)
        {
            string sql = "UPDATE Instrument SET QTY = @total, BI = @BI where PN = @PN";
            return connPNINOUT.Execute(sql, list.list);
        }

        [HttpPost]
        public int Update750System(SysList list)
        {
            string sql = "UPDATE J750_Board SET QTY = @total, BI = @BI where PN = @PN";
            return connPNINOUT.Execute(sql, list.list);
        }

        [HttpGet]
        public RespBI getBurnIn()
        {
            string sql = "SELECT PN, SN FROM [192.168.163.1].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            RespBI resp = new RespBI();
            List<BI> list = connPNINOUT.Query<BI>(sql).ToList();
            resp.List = list;
            resp.Message = "success";
            return resp;
        }



        [HttpGet]
        public List<Ins> GetInsList()
        {
            string sql = "select * from Instrument order by TYPE DESC, PN";
            return connPNINOUT.Query<Ins>(sql).ToList();
        }

        [HttpPost]
        public int AddIns(Ins ins)
        {
            string sql = @"Insert into Instrument
                           (PN, DES, MFR, OTC, TYPE, DAY, MOQ, ARQ, QTY, BI, Remark) values
                           (@PN, @DES, @MFR, @OTC, @TYPE, @DAY, @MOQ, @ARQ, @QTY, @BI, @Remark)";
            return connPNINOUT.Execute(sql, ins);
        }

        [HttpGet]
        public int DeleteIns(int id)
        {
            string sql = "delete from Instrument where ID = " + id;
            return connPNINOUT.Execute(sql);
        }

        [HttpPost]
        public int UpdateIns(Ins ins)
        {
            string sql = @"update Instrument set 
                           PN = @PN,
                           DES = @DES,
                           MFR = @MFR,
                           OTC = @OTC,
                           TYPE = @TYPE,
                           DAY = @DAY,
                           MOQ = @MOQ,
                           ARQ = @ARQ,
                           QTY = @QTY,
                           BI = @BI,
                           Remark = @Remark where ID = @ID";
            if (connPNINOUT.Execute(sql, ins) == 0)
            {
                return AddIns(ins);
            }
            return connPNINOUT.Execute(sql, ins);
        }

        private string OpenQueryFormat(string s)
        {
            return s.Replace("'", "''");
        }

        [HttpGet]
        public int RefreshQTY()
        {
            string sqlIns = "select * from Instrument order by TYPE DESC, PN";
            List<Ins> ins = connPNINOUT.Query<Ins>(sqlIns).ToList();
            string sqlRack = "select * from Rack";
            List<RackCar> racks = connPNINOUT.Query<RackCar>(sqlRack).ToList();
            string sqlBI = "select * FROM [192.168.163.1].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            List<BI> bis = connPNINOUT.Query<BI>(sqlBI).ToList();
            string pnList = string.Empty;
            StringBuilder sb = new StringBuilder();
            foreach (Ins i in ins)
            {
                sb.Append("'TDN-" + i.PN + "',");
            }
            pnList = sb.ToString();
            pnList = pnList.Substring(0, pnList.Length - 1);
            string sql = $@"SELECT cwar Warehouse,
                                   trim(item) ItemNo,
                                   stoc Qty
                            FROM bo_read011.whwmd215
                            WHERE stoc<>0
                                AND substr(item, 10, 3) IN ('TDN','MTS')
                                AND trim(item) IN ({pnList})";
            string openQuery = $"SELECT * FROM OPENQUERY(AS3P1, '{OpenQueryFormat(sql)}')";

            List<TdnOh> data = connBaan.Query<TdnOh>(openQuery).ToList();

            foreach (Ins i in ins)
            {
                i.QTY = 0;
                i.BI = 0;

                // P part 抓取Baan J11101 库位
                if (i.DES != null && i.DES.ToUpper().Contains("(P)"))
                {
                    foreach (TdnOh t in data)
                    {
                        if (t.ItemNo.Contains(i.PN) && t.Warehouse.Contains("101"))
                        {
                            i.QTY += t.Qty;
                        }
                    }
                }

                // DIB 抓取Baan J11203 库位
                else if (i.DES != null && i.DES.ToUpper().Contains("DIB"))
                {
                    foreach (TdnOh t in data)
                    {
                        if (t.ItemNo.Contains(i.PN) && t.Warehouse.Contains("203"))
                        {
                            i.QTY += t.Qty;
                        }
                    }
                }

                // 其余的都算PCBA板子 正常计算Supermarket和BI即可
                else
                {
                    foreach (RackCar rack in racks)
                    {
                        if (rack.PN.Contains("-"))
                        {
                            if (i.PN == rack.PN)
                            {
                                i.QTY += rack.ActualQTY;
                            }
                        }
                        else
                        {
                            string[] strs = rack.Command.Split(new char[] { ',', ' ', '"' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string str in strs)
                            {
                                if (i.PN == str)
                                {
                                    i.QTY++;
                                }
                            }
                        }
                    }

                    foreach (BI b in bis)
                    {
                        if (i.PN == b.PN)
                        {
                            i.BI++;
                        }
                    }
                }
            }

            string updateSql = "update Instrument set QTY = @QTY, BI = @BI WHERE ID = @ID";
            return connPNINOUT.Execute(updateSql, ins);
        }

        [HttpGet]
        public int JRefreshQTY()
        {
            string sqlIns = "select * from J750_Board order by PN";
            List<Ins> ins = connPNINOUT.Query<Ins>(sqlIns).ToList();
            string sqlRack = "select * from Rack";
            List<RackCar> racks = connPNINOUT.Query<RackCar>(sqlRack).ToList();

            foreach (Ins i in ins)
            {
                i.QTY = 0;
                i.BI = 0;
                
                foreach (RackCar rack in racks)
                {
                    if (rack.PN.Contains("-"))
                    {
                        if (i.PN == rack.PN)
                        {
                            i.QTY += rack.ActualQTY;
                        }
                    }
                    else
                    {
                        string[] strs = rack.Command.Split(new char[] { ',', ' ', '"' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string str in strs)
                        {
                            if (i.PN == str)
                            {
                                i.QTY++;
                            }
                        }
                    }
                }
            }

            string updateSql = "update J750_Board set QTY = @QTY WHERE ID = @ID";
            return connPNINOUT.Execute(updateSql, ins);
        }
    }
}