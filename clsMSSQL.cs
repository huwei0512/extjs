using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Data.SqlClient;
using System.IO;

namespace clsMSSQL
{
    public class MsSql
    {
        private string connectString;
        public string sErr = "";

        public MsSql(string connString)
        {
            //Persist Security Info=False;User ID=sa;pwd=Seu040991;Initial Catalog=CMS_Mirror;Data Source=10.137.12.32
            connectString = connString;
        }

        /// <summary>
        /// 查询数据库，并返回DataSet集
        /// </summary>
        /// <param name="cmdString">查询语句</param>
        /// <returns>返回DataSet</returns>
        public DataSet GetDataSet(string cmdString)
        {
            SqlConnection sqlConn = new SqlConnection(connectString);

            SqlDataAdapter da = new SqlDataAdapter(cmdString, sqlConn);
            DataSet ds = new DataSet();
            try
            {
                sqlConn.Open();

                da.Fill(ds);
                //if (ds != null || ds.Tables.Count > 0)
            }
            catch (SqlException sqlex)
            {
                Console.WriteLine(sqlex.Message);
                sErr = sqlex.Message;
            }
            sqlConn.Close();            
            return ds;
        }

        /// <summary>
        /// 取得表
        /// </summary>
        /// <param name="cmdString">查询字符串</param>
        /// <returns>DataTable</returns>
        public DataTable GetDataTable(string cmdString)
        {
            try
            {
                SqlConnection sqlConn = new SqlConnection(connectString);

                SqlDataAdapter da = new SqlDataAdapter(cmdString, sqlConn);
                DataSet ds = new DataSet();
                try
                {
                    sqlConn.Open();
                    da.Fill(ds);
                }
                catch (SqlException sqlex)
                {
                    Console.WriteLine(sqlex.Message);
                    sErr = sqlex.Message;
                }
                DataTable MyDataTable = ds.Tables[0];
                sqlConn.Close();
                return MyDataTable;
            }
            catch (Exception ex)
            {
                FCPortal.data.LogFile(cmdString + "@" + ex.Message, "sqlErr.txt");
                return null;
                //throw;
            }
        }

        /// <summary>
        /// 执行 sql语句
        /// </summary>
        /// <param name="cmdString">查询，更新，删除语句</param>
        /// <returns>成功返回true,失败返回false</returns>
        public bool ExecuteQuery(string cmdString)
        {
            SqlConnection conn = new SqlConnection(connectString);

            SqlCommand cmd = new SqlCommand(cmdString, conn);
            cmd.CommandText = cmdString;
            try
            {
                conn.Open();

                if (cmd.ExecuteNonQuery() > 0)
                {
                    conn.Close();
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (SqlException Sqlex)
            {
                //Console.WriteLine(Sqlex.Message);
                FCPortal.data.LogFile(cmdString+"@"+Sqlex.Message, "sqlErr.txt");
                sErr = Sqlex.Message;                
                return false;
            }
        }
                

    }

}
