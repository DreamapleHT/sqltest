using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data;
using DbAccess;
using System.Data.SqlClient;

namespace IBEcon.sys
{
    /// <summary>
    /// impExcel 的摘要说明
    /// </summary>
    public class impExcel : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            //参数1：opt
            string opt = "";
            context = context;
            if (context.Request["opt"] != null) opt = context.Request["opt"].ToString();
            else { opt = "students"; }
            switch (opt)
            {
                case "students": students(); break;
                case "hotele": hotele(); break;

            }
        }
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

        /// <summary>
        /// students方法
        /// </summary>
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
        /// <summary>
        /// 字段-字符串数组
        /// </summary>
        private string[] hotelColums ={
         "he_no","he_cname","he_ename"
         ,"he_sex","he_birth","he_nation","he_certtype"
         ,"he_certno","he_regtime","he_vautime","he_tempregion"
         ,"he_city","he_errtype","he_name","he_branch","he_police"
         ,"he_source","he_id"
                                     };
        /// <summary>
        /// s3t方法
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
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

        /// <summary>
        /// hotele方法
        /// </summary>
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

        private HttpContext context;
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}