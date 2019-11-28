using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ApiForControlTower.Models;
using Dapper;

namespace ApiForControlTower.Controllers
{
    public class BoardsController : ApiController
    {
        public IDbConnection connMinione = new SqlConnection(ConfigurationManager.ConnectionStrings["Minione"].ConnectionString);
        public IDbConnection connSupermarket = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        public List<BoardLocation> GetData()
        {
            // PN List
            string sql_Ins = "select * from [172.21.194.214].[PCBA].[dbo].[PNDes] order by PN";
            List<BoardLocation> BoardList = connMinione.Query<BoardLocation>(sql_Ins).ToList();

            // Supermarket
            string sql_Rack = "select * from [PNINOUT].[dbo].[Rack]";
            List<RackDetail> RackList = connSupermarket.Query<RackDetail>(sql_Rack).ToList();

            // Test
            string sql_Test_System = "select * from [172.21.194.214].[PCBA].[dbo].[LED]";
            List<TestSystem> TestSystemList = connMinione.Query<TestSystem>(sql_Test_System).ToList();
            string testSysStr = "";
            foreach (TestSystem testSystem in TestSystemList)
            {
                if (testSystem.System != "")
                {
                    if (testSysStr == "")
                    {
                        testSysStr = testSysStr + "'" + testSystem.System + "'";
                    }
                    else
                    {
                        testSysStr = testSysStr + ", '" + testSystem.System + "'";
                    }
                }
            }
            string sql_Test_Board = "select * from [172.21.194.214].[PCBA].[dbo].[Boards] where TesterID in (" + testSysStr + ")";
            List<TestBoard> TestBoardList = connMinione.Query<TestBoard>(sql_Test_Board).ToList();

            // Verify
            string sql_Verify_Board = "select * from [172.21.194.214].[PCBA].[dbo].[Boards] where Location = 'Verifying'";
            List<VerifyBoard> VerifyBoardList = connMinione.Query<VerifyBoard>(sql_Verify_Board).ToList();

            // Burn In
            string sql_BI_Board = "SELECT * FROM [172.21.194.214].[PCBA].[dbo].[Boards] where(type = 'Ultra' and location in ('supermarket', 'leanprocess', 'verifying')) or (type = 'flex' and location in ('supermarket', 'preburnin', 'verifying'))";
            List<BurnInBoard> BIBoardList = connMinione.Query<BurnInBoard>(sql_BI_Board).ToList();


            foreach (BoardLocation boardLocation in BoardList)
            {
                // Supermarket 统计
                boardLocation.SupermarketList = new List<SupermarketBoard>();
                foreach (RackDetail rack in RackList)
                {
                    if (rack.PN.Contains("Fixture") || rack.PN.Contains("Refurbish"))
                    {
                        for (int i = 0; i < rack.ActualQTY; i++)
                        {
                            SupermarketBoard sb = new SupermarketBoard();
                            sb.PN = rack.Command.Split(',')[i].Trim().Substring(1, rack.Command.Split(',')[i].Trim().Length - 2);
                            sb.SN = rack.SNView.Split(',')[i].Trim().Substring(1, rack.SNView.Split(',')[i].Trim().Length - 2);
                            sb.Location = rack.SlotView.Split(',')[i].Trim().Substring(1, rack.SlotView.Split(',')[i].Trim().Length - 2);
                            if (sb.PN == boardLocation.PN)
                            {
                                boardLocation.SupermarketList.Add(sb);
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < rack.ActualQTY; i++)
                        {
                            SupermarketBoard sb = new SupermarketBoard();
                            sb.PN = rack.PN;
                            sb.SN = rack.SNView.Split(',')[i].Trim().Substring(1, rack.SNView.Split(',')[i].Trim().Length - 2);
                            sb.Location = rack.SlotView.Split(',')[i].Trim().Substring(1, rack.SlotView.Split(',')[i].Trim().Length - 2);
                            if (sb.PN == boardLocation.PN)
                            {
                                boardLocation.SupermarketList.Add(sb);
                            }
                        }
                    }
                }

                // Test 统计
                boardLocation.TestList = new List<TestBoard>();
                foreach (TestBoard testBoard in TestBoardList)
                {
                    if (testBoard.PN == boardLocation.PN)
                    {
                        boardLocation.TestList.Add(testBoard);
                    }
                }

                // Verify 统计
                boardLocation.VerifyList = new List<VerifyBoard>();
                foreach (VerifyBoard verifyBoard in VerifyBoardList)
                {
                    if (verifyBoard.PN == boardLocation.PN)
                    {
                        boardLocation.VerifyList.Add(verifyBoard);
                    }
                }

                // Burn In 统计
                boardLocation.BurnInList = new List<BurnInBoard>();
                foreach (BurnInBoard BIBoard in BIBoardList)
                {
                    if (BIBoard.PN == boardLocation.PN)
                    {
                        boardLocation.BurnInList.Add(BIBoard);
                    }
                }
                
                boardLocation.SupermarketListQTY = boardLocation.SupermarketList.Count;
                boardLocation.TestListQTY = boardLocation.TestList.Count;
                boardLocation.VerifyListQTY = boardLocation.VerifyList.Count;
                boardLocation.BurnInListQTY = boardLocation.BurnInList.Count;
            }
            return BoardList;
        }
    }
}
