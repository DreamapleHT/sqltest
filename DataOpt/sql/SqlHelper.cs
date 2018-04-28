using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
namespace DataOpt.sql
{

    /// <summary>
    /// SQLServer操作类
    /// </summary>
    public class SqlHelper
    {
        //构造函数
        public SqlHelper()
        {
        }

        /// <summary>
        ///获取数据库连接字符串
        ///引用了System.Configuration
        /// </summary>
        public static string connectString = System.Configuration.ConfigurationManager.ConnectionStrings["SQLConntionString"].ConnectionString;


        #region 查询--无参数
        /// <summary>
        /// 查询方法--无参数
        /// </summary>
        /// <param name="cmdText">sql命令</param>
        /// <param name="cmdType">命令类型</param>
        /// <returns>返回json数据</returns>
        public static string QryData(string cmdText, CommandType cmdType)
        {
            //实例化数据库连接
            SqlConnection con = new SqlConnection(connectString);
            try
            {
                //打开连接
                con.Open();
                //适配器
                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                //命令类型
                da.SelectCommand.CommandType = cmdType;
                //实例化数据集
                DataSet ds = new DataSet();
                //填充到数据集
                da.Fill(ds, "data");
                //转换为json数据
                string dsString = ToolsFunction.dataSetToJson(ds);
                //返回数据集
                return dsString;
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                //关闭数据库连接
                con.Close();
            }


        }
        #endregion

        #region 查询--有参数
        /// <summary>
        /// 查询方法--有参数
        /// </summary>
        /// <param name="cmdText">sql命令</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="pars">参数数组</param>
        /// <returns>返回json数据</returns>
        public static string QryData(string cmdText, CommandType cmdType, params SqlParameter[] pars)
        {
            //实例化数据库连接
            SqlConnection con = new SqlConnection(connectString);
            try
            {
                //打开连接
                con.Open();

                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                //命令类型
                da.SelectCommand.CommandType = cmdType;
                //组合参数
                if (pars != null && pars.Length > 0)
                {
                    foreach (SqlParameter p in pars)
                    {
                        da.SelectCommand.Parameters.Add(p);
                    }
                }
                //实例化数据集
                DataSet ds = new DataSet();
                //填充到数据集
                da.Fill(ds, "data");
                //转换成json数据
                string dsString = ToolsFunction.dataSetToJson(ds);
                //返回数据集
                return dsString;
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                //关闭数据库连接
                con.Close();
            }


        }
        #endregion

        #region 查询--有参数
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cmdText">sql命令</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="pars">参数数组</param>
        /// <param name="type">默认参数，区分返回值,可传递任何值</param>
        /// <returns>返回DataSet数据集</returns>
        public static DataSet QryData(string cmdText, CommandType cmdType, string type = null, params SqlParameter[] pars)
        {
            //实例化数据库连接
            SqlConnection con = new SqlConnection(connectString);
            //打开连接
            con.Open();
            //适配器
            SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
            //命令类型
            da.SelectCommand.CommandType = cmdType;
            //获取并添加参数
            if (pars != null && pars.Length > 0)
            {
                foreach (SqlParameter p in pars)
                {
                    da.SelectCommand.Parameters.Add(p);
                }
            }
            //实例化数据集
            DataSet ds = new DataSet();
            //填充到数据集
            da.Fill(ds, "data");
            //关闭数据库连接
            con.Close();
            //返回数据集
            return ds;

        }
        #endregion


        #region 增删改--无参数
        /// <summary>
        /// sql的增删改操作
        /// </summary>
        /// <param name="cmdText">sql语句</param>
        /// <param name="cmdType">命令类型</param>
        /// <returns>返回影响行数</returns>
        public static int sqlOpt(string cmdText, CommandType cmdType)
        {

            SqlConnection con = new SqlConnection(connectString);

            con.Open();

            SqlCommand cmd = new SqlCommand(cmdText, con);

            cmd.CommandType = cmdType;

            int rows = cmd.ExecuteNonQuery();

            con.Close();

            return rows;

        }
        #endregion

        #region 增删改--有参数
        /// <summary>
        /// sql的增删改操作
        /// </summary>
        /// <param name="cmdText">sql语句</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="pars">命令参数</param>
        /// <returns>返回影响行数</returns>
        public static int sqlOpt(string cmdText, CommandType cmdType, params SqlParameter[] pars)
        {

            SqlConnection con = new SqlConnection(connectString);
            try
            {
                con.Open();

                SqlCommand cmd = new SqlCommand(cmdText, con);

                cmd.CommandType = cmdType;

                if (pars != null && pars.Length > 0)
                {

                    foreach (SqlParameter p in pars)
                    {
                        cmd.Parameters.Add(p);

                    }

                }

                int rows = cmd.ExecuteNonQuery();

                return rows;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                con.Close();
            }

        }
        #endregion

        #region 执行存储过程-无参数
        /// <summary>
        /// 存储过程
        /// </summary>
        /// <param name="procName">存储过程名称</param>
        /// <returns></returns>
        public static string SqlProc(string procName)
        {
            SqlConnection con = new SqlConnection(connectString);
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter();

            da.SelectCommand = new SqlCommand(procName, con);
            //存储过程
            da.SelectCommand.CommandType = CommandType.StoredProcedure;
            //SqlParameter id = da.SelectCommand.Parameters.Add("@id", SqlDbType.NText);
            DataSet ds = new DataSet();
            da.Fill(ds, "data");
            con.Close();
            string dsString = ToolsFunction.dataSetToJson(ds);
            return dsString;
        }
        #endregion

        #region 执行存储过程-有参数
        /// <summary>
        /// 存储过程
        /// </summary>
        /// <param name="procName">存储过程名称</param>
        /// <param name="pars">参数</param>
        /// <returns></returns>
        public static string SqlProc(string procName, params SqlParameter[] pars)
        {
            SqlConnection con = new SqlConnection(connectString);
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter();

            da.SelectCommand = new SqlCommand(procName, con);
            //存储过程
            da.SelectCommand.CommandType = CommandType.StoredProcedure;

            //设置存储过程的参数值,其中@id 为存储过程的参数.
            if (pars != null && pars.Length > 0)
            {

                foreach (SqlParameter p in pars)
                {

                    da.SelectCommand.Parameters.Add(p);

                }

            }
            //SqlParameter id = da.SelectCommand.Parameters.Add("@id", SqlDbType.NText);
            DataSet ds = new DataSet();
            da.Fill(ds, "data");
            con.Close();
            string dsString = ToolsFunction.dataSetToJson(ds);
            return dsString;
        }
        #endregion

        #region 无返回值--事务--统一返回值类型--统一入口

        #endregion

    }

}
