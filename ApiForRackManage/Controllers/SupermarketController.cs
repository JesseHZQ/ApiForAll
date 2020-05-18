using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForRackManage.Controllers
{
    public class SupermarketController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        public List<RackInfo> GetRacks()
        {
            string sql = "select * from Supermarket_Rack";
            return conn.Query<RackInfo>(sql).ToList();
        }

        public List<Board> GetBoards ()
        {
            List<Board> BoardList = new List<Board>();
            string sql = "select * from Supermarket_Rack";
            List<Ra> RackList = conn.Query<Ra>(sql).ToList();

            string sqlRefurbish = "select * from Supermarket_Refurbish";
            List<Refurbish> RefList = conn.Query<Refurbish>(sqlRefurbish).ToList();

            foreach (Ra rack in RackList)
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
                    Refurbish r = RefList.Where(x => x.SN.Trim().ToUpper() == board.SN.Trim().ToUpper()).FirstOrDefault();
                    board.IsRefurbish = (r != null);
                    BoardList.Add(board);
                }
            }
            return BoardList.Where(x => x.IsRefurbish).ToList();
        }

        private class Ra
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

        private class Refurbish
        {
            public int Id { get; set; }
            public string SN { get; set; }
        }
    }
}
