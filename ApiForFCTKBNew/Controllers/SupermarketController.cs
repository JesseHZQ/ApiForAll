using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;

namespace ApiForFCTKBNew.Controllers
{
    public class SupermarketController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);
        public IDbConnection connPNINOUT = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);
        public IDbConnection connBaan = new SqlConnection(ConfigurationManager.ConnectionStrings["Baan"].ConnectionString);

        [HttpGet]
        public int RefreshQTY()
        {
            string sqlIns = "select * from Instrument order by TYPE DESC, PN";
            List<Ins> ins = connPNINOUT.Query<Ins>(sqlIns).ToList();
            List<Board> boardList = GetAllBoards();
            //string sqlRack = "select * from Supermarket_Rack";
            //List<RackCar> racks = connPNINOUT.Query<RackCar>(sqlRack).ToList();
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
                else if (i.DES != null && (i.DES.ToUpper().Contains("DIB") || i.DES.ToUpper().Contains("DB")))
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
                    i.QTY = boardList.Where(x => x.PN == i.PN).ToList().Count;

                    //foreach (RackCar rack in racks)
                    //{
                    //    if (rack.PN.Contains("-"))
                    //    {
                    //        if (i.PN == rack.PN)
                    //        {
                    //            i.QTY += rack.ActualQTY;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        string[] strs = rack.Command.Split(new char[] { ',', ' ', '"' }, StringSplitOptions.RemoveEmptyEntries);
                    //        foreach (string str in strs)
                    //        {
                    //            if (i.PN == str)
                    //            {
                    //                i.QTY++;
                    //            }
                    //        }
                    //    }
                    //}

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

        
        public List<Board> GetAllBoards()
        {
            List<Board> BoardList = new List<Board>();
            string sql = "select * from Supermarket_Rack";
            List<RackCar> RackList = connPNINOUT.Query<RackCar>(sql).ToList();
            foreach (RackCar rack in RackList)
            {
                List<string> PNList = rack.PNList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> SNList = rack.SNList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> SlotList = rack.SlotList.Split(new char[] { '"', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                List<string> DateList = rack.DateList.Split(new char[] { '"', ',' }, StringSplitOptions.RemoveEmptyEntries).Where(x => x != " ").ToList();
                for (int i = 0; i < PNList.Count; i++)
                {
                    Board board = new Board();
                    board.PN = PNList[i];
                    board.SN = SNList[i];
                    board.RackName = rack.RackName;
                    board.Slot = SlotList[i];
                    board.Date = DateTime.Parse(DateList[i]);
                    BoardList.Add(board);
                }
            }
            return BoardList;
        }


        private string OpenQueryFormat(string s)
        {
            return s.Replace("'", "''");
        }
    }

    public class RackCar
    {
        public int RackId { get; set; }
        public string RackName { get; set; }
        public int RackSize { get; set; }
        public int TypeId { get; set; }
        public string PN { get; set; }
        public string PNList { get; set; }
        public string SNList { get; set; }
        public string SlotList { get; set; }
        public string DateList { get; set; }
    }

    internal class TdnOh
    {
        public string Warehouse { get; set; }
        public string ItemNo { get; set; }
        public int Qty { get; set; }
    }

    internal class BI
    {
        public string PN { get; set; }
        public string SN { get; set; }
    }

    public class Board
    {
        public string PN { get; set; }
        public string SN { get; set; }
        public string RackName { get; set; }
        public string Slot { get; set; }
        public DateTime Date { get; set; }
        public bool IsRefurbish { get; set; }
    }

    internal class Ins
    {
        public int ID { get; set; }
        public string PN { get; set; }
        public string DES { get; set; }
        public string MFR { get; set; }
        public string OTC { get; set; }
        public string TYPE { get; set; }
        public int DAY { get; set; }
        public int MOQ { get; set; }
        public int ARQ { get; set; }
        public int QTY { get; set; }
        public int BI { get; set; }
        public int WH { get; set; }
        public string Remark { get; set; }
    }
}
