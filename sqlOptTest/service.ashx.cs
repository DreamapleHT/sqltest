using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataOpt;

using System.Data;
//using System.Text;
using DataOpt.sql;
using System.Collections;
using System.Data.SqlClient;
namespace sqlOptTest
{
    /// <summary>
    /// login 的摘要说明
    /// </summary>
    public class login : IHttpHandler
    {
        #region 对请求的处理
        public void ProcessRequest(HttpContext context)
        {
            //请求类型
            string rt = context.Request.RequestType;

            string sqlcmd = "";
            string procName = "";
            string optType = "";
            string mobile = "";
            string account = "";
            string id = "";

            //get和post请求通用的方法
            //context.Response.ContentType = "text/plain";
            //procName = context.Request["procName"];
            //idvalue = context.Request["idvalue"];
            //optType = context.Request["optType"];
            if (rt == "GET")
            {
                //仅限get
                context.Response.ContentType = "text/plain";
                procName = context.Request.QueryString["procName"];
                optType = context.Request.QueryString["optType"];
                mobile = context.Request.QueryString["mobile"];
                account = context.Request.QueryString["account"];
                id = context.Request.QueryString["id"];


            }
            else
            {
                //仅限post
                context.Response.ContentType = "text/plain";
                procName = context.Request.Form["procName"];
                optType = context.Request.Form["optType"];
                mobile = context.Request.Form["mobile"];
                account = context.Request.Form["account"];
                id = context.Request.Form["id"];


            }

            //判断操作类型
            switch (optType)
            {
                case "qry":
                    {
                        SqlParameter[] parameters = { 
                                         new SqlParameter("@mobile",mobile)  
                                        };
                        sqlcmd = "select * from sys_user where su_mobile=@mobile";
                        string backData = SqlHelper.QryData(sqlcmd, CommandType.Text, parameters);
                        context.Response.Write(backData);
                        break;
                    }
                case "ins":
                    {
                        if (account == "" || mobile == "")
                        {
                            context.Response.Write("账户和手机号码不能为空");
                            return;
                        }
                        SqlParameter[] parameters = { 
                                        new SqlParameter("@account",account),
                                         new SqlParameter("@mobile",mobile)  
                                        };
                        sqlcmd = "INSERT INTO sys_user(su_account,su_mobile) VALUES( @account  ,@mobile )";
                        int backRows = SqlHelper.sqlOpt(sqlcmd, CommandType.Text, parameters);
                        context.Response.Write(backRows);
                        break;
                    }
                case "edt":
                    {
                        SqlParameter[] parameters = { 
                                        new SqlParameter("@account",account),
                                         new SqlParameter("@mobile",mobile)  
                                        };
                        sqlcmd = "UPDATE sys_user SET su_account =@account where su_mobile=@mobile";
                        int backRows = SqlHelper.sqlOpt(sqlcmd, CommandType.Text, parameters);

                        context.Response.Write(backRows);
                        break;
                    }
                case "del":
                    {
                        SqlParameter[] parameters = { 
                                      
                                        };
                        //
                        sqlcmd = "delete from sys_user";
                        int backRows = SqlHelper.sqlOpt(sqlcmd, CommandType.Text, parameters);
                        context.Response.Write(backRows);
                        break;
                    }
                case "proc":
                    {
                        SqlParameter[] parameters = { 
                                        new SqlParameter("@id",id),
                                        };
                        string backData = SqlHelper.SqlProc(procName, parameters);
                        context.Response.Write(backData);
                        break;
                    }
            }

        }

        #endregion

        //获取一个值，该值指示其他请求是否可以使用 IHttpHandler 实例。
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

    }
}