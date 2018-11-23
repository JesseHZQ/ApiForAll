using System;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace GROUP.Manage
{
	/// <summary>
	/// BaseClass 的摘要说明。
	/// </summary>
	public class BaseClass: System.Web.UI.Page
	{
        String strConn;
        public BaseClass()
		{
            strConn = ConfigurationManager.ConnectionStrings["Minione"].ConnectionString;
            
		}

		//读写数据表--DataTable
		public DataTable ReadTable(String strSql)
		{

            //DataTable dt = new DataTable();//创建一个数据表dt
            //SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化

            //SqlTransaction myTrans = Conn.BeginTransaction(); //使用New新生成一个事务
            //SqlCommand myCommand = new SqlCommand();
            //myCommand.Transaction = myTrans;
            //try
            //{


            //    Conn.Open();//打开连接
            //    SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            //    Cmd.Fill(dt);                               //将数据适配器中的数据填充到数据集dt中
            //    Conn.Close();//关闭连接
            //    return dt;

            //    Console.WriteLine("Record is updated.");
            //}
            //catch (Exception e)
            //{
            //    myTrans.Rollback();
            //    Console.WriteLine(e.ToString());
            //    Console.WriteLine("Sorry, Record can not be updated.");
            //}
            //finally
            //{
            //    Conn.Close();
            //}
            //return dt;

            DataTable dt = new DataTable();//创建一个数据表dt
            SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化
            Conn.Open();//打开连接
            SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            Cmd.Fill(dt);								//将数据适配器中的数据填充到数据集dt中
            Conn.Close();//关闭连接
            return dt;
        }

        public DataTable ReadTable_Other_DataBase(String strSql, String strConnection)
        {
            DataTable dt = new DataTable();//创建一个数据表dt
            SqlConnection Conn = new SqlConnection(strConnection);//定义新的数据连接控件并初始化
            Conn.Open();//打开连接
            SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            Cmd.Fill(dt);                               //将数据适配器中的数据填充到数据集dt中
            Conn.Close();//关闭连接
            return dt;
        }

        //读写数据集--DataSet
        public DataSet ReadDataSet(String strSql)
		{
			DataSet ds=new DataSet();//创建一个数据集ds
            SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化
            Conn.Open();//打开连接
            SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            Cmd.Fill(ds);								//将数据适配器中的数据填充到数据集ds中
			Conn.Close();//关闭连接
			return ds;
		}

        public DataSet GetDataSet(String strSql, String tableName)
		{
            DataSet ds = new DataSet();//创建一个数据集ds
            SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化
            Conn.Open();//打开连接
            SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            Cmd.Fill(ds, tableName);								//将数据适配器中的数据填充到数据集ds中
            Conn.Close();//关闭连接
            return ds;

		}

        public SqlDataReader readrow(String sql)
		{

            SqlConnection Conn = new SqlConnection(strConn);
            Conn.Open();

            SqlCommand Comm = new SqlCommand(sql, Conn);
            SqlDataReader Reader = Comm.ExecuteReader();


            if (Reader.Read())
			{
                Comm.Dispose();
                return Reader;
			}
			else 
			{
                Comm.Dispose();
				return null;
			}

		}

		//读某一行中某一字段的值
        public string Readstr(String strSql, int flag)
		{
			DataSet ds=new DataSet();//创建一个数据集ds
            String str;

            SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化

            Conn.Open();//打开连接

            SqlDataAdapter Cmd = new SqlDataAdapter(strSql, Conn);//定义并初始化数据适配器
            Cmd.Fill(ds);								//将数据适配器中的数据填充到数据集ds中

			str=ds.Tables[0].Rows[0].ItemArray[flag].ToString();
			Conn.Close();//关闭连接
			return str;
		}


        public void execsql(String strSql)
		{
           // SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化

           // SqlCommand Comm = new SqlCommand(strSql, Conn);

           // Conn.Open();

           //SqlTransaction myTrans = Conn.BeginTransaction(); //使用New新生成一个事务

           // Comm.Transaction = myTrans;
           // try
           // {
              

           //     Comm.ExecuteNonQuery();

           //     Console.WriteLine("Record is updated.");
           // }
           // catch (Exception e)
           // {
           //     myTrans.Rollback();
           //     Console.WriteLine(e.ToString());
           //     Console.WriteLine("Sorry, Record can not be updated.");
           // }
           // finally
           // {
           //     Conn.Close();
           // }
            SqlConnection Conn = new SqlConnection(strConn);//定义新的数据连接控件并初始化

            SqlCommand Comm = new SqlCommand(strSql, Conn);
            Conn.Open();//打开连接

            Comm.ExecuteNonQuery();//执行命令
            Conn.Close();//关闭连接

        }


        public SqlConnection createcon()
        {
            throw new NotImplementedException();
        }
    }
}
