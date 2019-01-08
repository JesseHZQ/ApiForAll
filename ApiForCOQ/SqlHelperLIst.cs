using System.Configuration;
using System.Data;
using System.Data.SqlClient;
namespace ApiForCOQ
{
    public class SqlHelperList
    {
        //连接字符串
        //1、添加引用 2、导入命名空间 为了使用ConfigurationManager
        public static string conStr = ConfigurationManager.ConnectionStrings["COQLIST"].ConnectionString;

        //增删改查
        //查找数据  ExecuteScalar()返回首行首列   ExecuteReader()  DataTable

        /// <summary>
        /// 返回DataTable
        /// </summary>
        /// <param name="sql">所用的sql语句</param>
        /// <param name="param">可变，可以传参也可以不传参数</param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, params SqlParameter[] param)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(conStr))
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter(sql, con))
                {
                    //添加参数
                    adapter.SelectCommand.Parameters.AddRange(param);
                    //1.打开链接，如果连接没有打开，则它给你打开；如果打开，就算了
                    //2.去执行sql语句，读取数据库
                    //3.sqlDataReader,把读取到的数据填充到内存表中
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        /// <summary>
        /// 执行查询，返回首行首列
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static object ExecuteScalar(string sql, params SqlParameter[] param)
        {
            object o = null;
            using (SqlConnection con = new SqlConnection(conStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Parameters.AddRange(param);
                    con.Open();

                    o = cmd.ExecuteScalar();
                }
            }
            return o;
        }


        /// <summary>
        /// 执行查询，返回SqlDataReader对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static SqlDataReader ExecuteReader(string sql, params SqlParameter[] param)
        {
            SqlDataReader reader = null;
            using (SqlConnection con = new SqlConnection(conStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Parameters.AddRange(param);
                    con.Open();

                    reader = cmd.ExecuteReader();
                }
            }
            return reader;
        }

        /// <summary>
        /// 执行增删改，返回受影响的行数
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(string sql, params SqlParameter[] param)
        {
            int n = -1;
            using (SqlConnection con = new SqlConnection(conStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Parameters.AddRange(param);
                    con.Open();
                    n = cmd.ExecuteNonQuery();
                }
            }
            return n;
        }

    }
}