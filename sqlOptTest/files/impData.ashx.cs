using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data;
using DbAccess;
using System.Data.SqlClient;
using System.Text;
using Aspose.Cells;
using DataOpt.sql;

namespace sqlOptTest.files
{
    /// <summary>
    /// impData 的摘要说明
    /// </summary>
    public class impData : IHttpHandler
    {
        private HttpContext context;
        //文件地址以及名称
        private string fileUrl;

        public void ProcessRequest(HttpContext context1)
        {
            context1.Response.ContentType = "text/plain";
            context1.Response.ContentEncoding = Encoding.UTF8;
            if (context1.Request["REQUEST_METHOD"] == "OPTIONS")
            {
                context1.Response.End();
            }

            //参数1：opt
            //string opt = "";
            //context1.Response.Write("text/plain");
            //context1.Response.Write(context.Request["path"]);
            context = context1;
            //if (context1.Request["opt"] != null) opt = context1.Request["opt"].ToString();
            //else { opt = "students"; }
            //switch (opt)
            //{
            //    case "students": students(); break;
            //    case "hotele": hotele(); break;

            //}
            //保存文件
            SaveFile();
            DataSet dsExcel = new DataSet();
            //读取数据到dataset
            dsExcel = impDataSin(fileUrl, "1");
            //存入数据库
            addData(dsExcel);

        }



        #region 文件保存
        /// <summary>
        /// 文件保存
        /// </summary>
        /// <param name="basePath">保存地址</param>
        private void SaveFile(string basePath = "~/attach/files/")
        {
            var name = string.Empty;
            basePath = (basePath.IndexOf("~") > -1) ? System.Web.HttpContext.Current.Server.MapPath(basePath) : basePath;
            HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
            //如果目录不存在，则创建目录
            if (!Directory.Exists(basePath))
            {
                Directory.CreateDirectory(basePath);
            }

            var suffix = files[0].ContentType.Split('/');
            //获取文件格式
            var _suffix = suffix[1].Equals("jpeg", StringComparison.CurrentCultureIgnoreCase) ? "" : suffix[1];
            var _temp = System.Web.HttpContext.Current.Request["name"];
            //如果不修改文件名，则创建随机文件名
            if (!string.IsNullOrEmpty(_temp))
            {
                name = _temp;
            }
            else
            {
                Random rand = new Random(24 * (int)DateTime.Now.Ticks);
                name = rand.Next() + "." + _suffix;
            }
            //文件保存
            var full = basePath + name;
            fileUrl = full;
            files[0].SaveAs(full);
            var _result = "{\"jsonrpc\" : \"2.0\", \"result\" : null, \"id\" : \"" + name + "\"}";
            System.Web.HttpContext.Current.Response.Write(_result);
        }
        #endregion


        #region 单文件读取》单表存储》从第二行读取数据
        public static DataSet impDataSin(string excelPath, string sheetName)
        {
            DataSet ds = new DataSet();
            Workbook workBook = new Workbook(excelPath);
            Worksheet workSheet = workBook.Worksheets[sheetName];
            Cells cell = workSheet.Cells;
            DataTable dt = new DataTable();
            int count = cell.Columns.Count;
            for (int i = 0; i < count; i++)
            {
                string str = cell.GetRow(0)[i].StringValue;
                dt.Columns.Add(new DataColumn(str));
            }
            for (int i = 1; i < cell.Rows.Count; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < count; j++)
                {
                    dr[j] = cell[i, j].StringValue;
                }
                dt.Rows.Add(dr);
            }
            dt.AcceptChanges();
            ds.Tables.Add(dt);
            return ds;
        }
        #endregion


        #region 存入数据库
        /// <summary>
        /// 数据入库
        /// </summary>
        /// <param name="data">数据集</param>
        /// <returns>返回影响行数</returns>
        private static int addData(DataSet data)
        {
            int rows = 0;
            //循环数据集
            foreach (DataRow da in data.Tables[0].Rows)
            {
                //sql模板语句
                string sql = string.Format("insert into sys_user (su_account) values ({0})", da["su_account"].ToString());
                //改变影响行数
                rows += SqlHelper.sqlOpt(sql, CommandType.Text);
            }
            return rows;
        }
        #endregion

        #region 单文件读取》多表存储
        #endregion

        /// <summary>
        /// 字段-字符串数组
        /// </summary>
        private string[] stuColums = 
        { "bs_name", "bs_sex", "bs_nation" 
        ,"bs_birth","bs_passno","bs_cardno"
        ,"bs_cardtime","bs_profess","bs_stime"
        ,"bs_type","bs_term","bs_addr","bs_tel"
        ,"bs_id","bs_createmen","bs_createtime"
        ,"bs_sclname"
        };

        ///// <summary>
        ///// students方法
        ///// </summary>
        public void students()
        {
            HttpPostedFile postedFile = context.Request.Files["Filedata"];
            string ext = context.Request["ext"].ToString();
            string author = "";
            string school = "";
            try
            {
                author = Functions.tostr(context.Request["ctemen"].ToString());
            }
            catch
            {
                author = "";
            }
            try
            {
                school = Functions.tostr(context.Request["school"].ToString());
            }
            catch
            {
                school = "";
            }
            string filepath = context.Server.MapPath("/");
            filepath += @"\upload\cache";
            if (!Directory.Exists(filepath)) Directory.CreateDirectory(filepath);
            filepath += @"\cache_" + DateTime.Now.Millisecond + "." + ext;
            SqlConnection conn = new SqlConnection(Functions.connstr());
            if (conn.State != ConnectionState.Open) conn.Open();
            try
            {
                postedFile.SaveAs(filepath);
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);
                Aspose.Cells.Cells cells = workbook.Worksheets[0].Cells;
                DataTable dt = cells.ExportDataTable(1, 0, cells.MaxRow, stuColums.Length);
                DataTable dt1 = new DataTable();
                for (int i = 0; i < stuColums.Length; i++)
                {
                    dt1.Columns.Add(stuColums[i]);
                }
                int counter = 0;
                for (int i = counter; i < dt.Rows.Count; i++)
                {
                    if ((dt.Rows[i][0] != null && dt.Rows[i][0].ToString() != "") && (dt.Rows[i][4] != null && dt.Rows[i][4].ToString() != ""))
                    {
                        dt1.Rows.Add(dt.Rows[i].ItemArray);
                        //设置自定义规则的id
                        dt1.Rows[i - counter]["bs_id"] = Functions.getprimarykey("ae_a.dbo.busi_students", conn);
                        dt1.Rows[i - counter]["bs_createmen"] = author == "" ? "系统导入" : (author + "-导入");
                        dt1.Rows[i - counter]["bs_createtime"] = DateTime.Now;
                        dt1.Rows[i - counter]["bs_sclname"] = school;
                    }
                    else
                    {
                        counter++;
                    }
                }
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(Functions.connstr()))
                {
                    bcp.DestinationTableName = "ae_a.dbo.busi_students";
                    for (int i = 0; i < stuColums.Length; i++)
                    {
                        bcp.ColumnMappings.Add(stuColums[i], stuColums[i]);
                    }
                    bcp.WriteToServer(dt1);
                }
                context.Response.Write("OK");
            }
            catch (Exception e)
            {
                context.Response.Write("-1," + e.Message);
            }
            File.Delete(filepath);
            context.Response.End();

        }
        ///// <summary>
        ///// 字段-字符串数组
        ///// </summary>
        private string[] hotelColums ={
         "he_no","he_cname","he_ename"
         ,"he_sex","he_birth","he_nation","he_certtype"
         ,"he_certno","he_regtime","he_vautime","he_tempregion"
         ,"he_city","he_errtype","he_name","he_branch","he_police"
         ,"he_source","he_id"
                                     };
        ///// <summary>
        ///// s3t方法
        ///// </summary>
        ///// <param name="s"></param>
        ///// <returns></returns>
        public string s2t(string s)
        {
            string res = "";
            if (s.Length > 8)
            {
                if (s.Length > 14) s = s.Substring(0, 14);
                if (s.Length == 14)
                {
                    res += s.Substring(0, 4);
                    res += "-" + s.Substring(4, 2);
                    res += "-" + s.Substring(6, 2);
                    res += " " + s.Substring(8, 2);
                    res += ":" + s.Substring(10, 2);
                    res += ":" + s.Substring(12, 2);
                    return res;
                }
                else
                {
                    s = s.Substring(0, 8);
                }
            }
            if (s.Length == 8)
            {
                res += s.Substring(0, 4);
                res += "-" + s.Substring(4, 2);
                res += "-" + s.Substring(6, 2);
                return res;
            }
            return res;
        }

        ///// <summary>
        ///// hotele方法
        ///// </summary>
        public void hotele()
        {
            HttpPostedFile postedFile = context.Request.Files["Filedata"];
            string ext = context.Request["ext"].ToString();
            string author = "";
            try
            {
                author = Functions.tostr(context.Request["ctemen"].ToString());
            }
            catch
            {
                author = "";
            }
            string filepath = context.Server.MapPath("/");
            filepath += @"\upload\cache";
            if (!Directory.Exists(filepath)) Directory.CreateDirectory(filepath);
            filepath += @"\cache_" + DateTime.Now.Millisecond + "." + ext;
            SqlConnection conn = new SqlConnection(Functions.connstr());
            if (conn.State != ConnectionState.Open) conn.Open();
            try
            {
                postedFile.SaveAs(filepath);
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filepath);
                Aspose.Cells.Cells cells = workbook.Worksheets[0].Cells;
                DataTable dt = cells.ExportDataTable(1, 0, cells.MaxRow, hotelColums.Length);
                DataTable dt1 = new DataTable();
                for (int i = 0; i < hotelColums.Length; i++)
                {
                    dt1.Columns.Add(hotelColums[i]);
                }
                int counter = 1;
                for (int i = counter; i < dt.Rows.Count; i++)
                {
                    if ((dt.Rows[i][0] != null && dt.Rows[i][0].ToString() != "") && (dt.Rows[i][13] != null && dt.Rows[i][13].ToString() != ""))
                    {
                        dt1.Rows.Add(dt.Rows[i].ItemArray);
                        dt1.Rows[i - counter]["he_birth"] = s2t(dt1.Rows[i - counter]["he_birth"].ToString());
                        dt1.Rows[i - counter]["he_regtime"] = s2t(dt1.Rows[i - counter]["he_regtime"].ToString());
                        dt1.Rows[i - counter]["he_vautime"] = s2t(dt1.Rows[i - counter]["he_vautime"].ToString());
                        dt1.Rows[i - counter]["he_id"] = Functions.getprimarykey("ae_a.dbo.hotel_eregis", conn);
                    }
                    else
                    {
                        counter++;
                    }
                }
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(Functions.connstr()))
                {
                    bcp.DestinationTableName = "ae_a.dbo.hotel_eregis";
                    for (int i = 0; i < hotelColums.Length; i++)
                    {
                        bcp.ColumnMappings.Add(hotelColums[i], hotelColums[i]);
                    }
                    bcp.WriteToServer(dt1);
                }
                context.Response.Write("OK");
            }
            catch (Exception e)
            {
                context.Response.Write("-1," + e.Message);
            }
            File.Delete(filepath);
            context.Response.End();
        }






        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}