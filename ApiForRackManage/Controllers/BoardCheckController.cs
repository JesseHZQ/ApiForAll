using ApiForRackManage.Models;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;

namespace ApiForRackManage.Controllers
{
    public class BoardCheckController : ApiController
    {
        public IDbConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["PNINOUT"].ConnectionString);

        /// <summary>
        /// 接收前端上传单个文件
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public List<ConfigBoard> UploadFile()
        {
            List<ConfigBoard> configBoards = new List<ConfigBoard>();
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile f = httpRequest.Files[0];
            if (!Directory.Exists(HttpContext.Current.Server.MapPath("../../Uploads")))
            {
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("../../Uploads"));
            }
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("../../Uploads"), Path.GetFileName(f.FileName));
            f.SaveAs(filePath);
            StringBuilder sb = new StringBuilder();
            List<string> strList = File.ReadAllLines(filePath).ToList();
            int begin = strList.FindIndex(x => x == "<TEST_HEAD TH 1>");
            int end = strList.FindIndex(x => x == "</TEST_HEAD TH 1>");
            strList = strList.GetRange(begin + 3, end - begin - 3);
            string slot = "";
            int index = 0;
            foreach (string str in strList)
            {
                ConfigBoard cb = new ConfigBoard();
                string[] strs = str.Split(new char[] { '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (strs.Length == 6)
                {
                    slot = strs[0];
                    index = 0;
                    cb.Slot = strs[0];
                    cb.Type = strs[1];
                    cb.PN = strs[2];
                    cb.SN = strs[3];
                    cb.RevD = strs[4].Split('-')[0];
                    cb.RevL = strs[4].Split('-')[1];
                }
                else
                {
                    index++;
                    cb.Slot = slot + '.' + index;
                    cb.Type = strs[0];
                    cb.PN = strs[1];
                    cb.SN = strs[2];
                    cb.RevD = strs[3].Split('-')[0];
                    cb.RevL = strs[3].Split('-')[1];
                }
                configBoards.Add(cb);
            }
            File.Delete(filePath);
            return configBoards;
        }

        [HttpGet]
        public List<ConfigBoard> GetConfigBoardsByMark(string mark)
        {
            string sql = "SELECT * FROM [192.168.163.1].[Boardscheck].[dbo].[Boardscheck] WHERE Mark = '" + mark + "'";
            return conn.Query<ConfigBoard>(sql).ToList();
        }

        [HttpPost]
        public int UploadData(List<ConfigBoard> list)
        {
            foreach (ConfigBoard item in list)
            {
                item.CurrentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            }
            if (list != null)
            {
                string sql = "";
                string s = "SELECT * FROM [192.168.163.1].[Boardscheck].[dbo].[Boardscheck] WHERE Mark = '" + list[0].Mark + "'";
                List<ConfigBoard> l = conn.Query<ConfigBoard>(s).ToList();
                if (l.Count == 0)
                {
                    sql = "INSERT INTO [192.168.163.1].[Boardscheck].[dbo].[Boardscheck] (Mark, PN, SN, RevD, RevL, ID, CurrentTime, item) VALUES (@Mark, @PN, @SN, @RevD, @RevL, @ID, @CurrentTime, @item)";
                }
                else
                {
                    sql = "UPDATE [192.168.163.1].[Boardscheck].[dbo].[Boardscheck] SET Mark = @Mark, PN = @PN, SN = @SN, RevD = @RevD, RevL = @RevL, ID = @ID, CurrentTime = @CurrentTime, item = @item WHERE Mark = @Mark AND PN = @PN";
                }
                return conn.Execute(sql, list);
            }
            else
            {
                return 0;
            }
            
            
        }

        [HttpPost]
        public User Login(User user)
        {
            string sql = "SELECT * FROM [192.168.163.1].[Boardscheck].[dbo].[Users] WHERE ID = @ID AND Password = @Password";
            return conn.QueryFirstOrDefault<User>(sql, user);
        }
    }
}
