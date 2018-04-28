using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Data.SqlClient;
using Aspose.Cells;
using System.IO;
using DbAccess;
using System.Data;
using System.Configuration;
using System;
using System.Text.RegularExpressions;


namespace IBComm.upload.export
{
    /// <summary>
    /// exp_metricsdata 的摘要说明
    /// </summary>
    public class exp_metricsdata : IHttpHandler
    {
        string success = "";
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            try
            {
                execsqldb(context);
            }
            catch (Exception ex)
            {
                success = "-1" + ex.ToString();
            }
            finally
            {
                context.Response.Write(success);
            }
        }

        private void execsqldb(HttpContext context)
        {

            using (SqlConnection conn = new SqlConnection(Functions.connstr()))
            {
                conn.Open();
                exporttoexcel(context, conn);
                conn.Close();
                conn.Dispose();
            }
        }

        private void exporttoexcel(HttpContext page, SqlConnection conn)
        {
            string filename = "";
            if (page.Request["filename"] != null)
            {
                filename = Functions.tostr(page.Request["filename"].ToString()).Trim();
            }

            string strtempfile = page.Server.MapPath("/") + @"upload\templet\" + filename + ".xls";
            Workbook workbook = new Workbook();
            workbook.Open(strtempfile);

            Style style1 = workbook.Styles[workbook.Styles.Add()];//新增样式
            style1.HorizontalAlignment = TextAlignmentType.Right;//文字居中
            style1.Font.Name = "宋体";//文字字体
            style1.Font.Size = 11;//文字大小

            style1.IsTextWrapped = true;//单元格内容自动换行
            style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线
            style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线
            style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线
            style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线;

            string exportype = Functions.tostr(page.Request["exportype"].ToString()).Trim();
            string metricsid = Functions.tostr(page.Request["metricsid"].ToString()).Trim();

            string sql = "";
            sql = " select b.c_uniqueid,b.c_name from mps_metricscla a, mps_metricscla b where a.c_uniqueid = '" + metricsid + "' " +
            "and b.c_id like a.c_id + '%' and b.c_uniqueid in (select mi_metricsid from mps_metricsitem )";
           
            DataSet ds1 = loadfiexddata(conn, sql);
            string cname = "";
            string uniqueid = "";
            foreach (DataRow drow2 in ds1.Tables[0].Rows)
            {
                uniqueid = drow2["c_uniqueid"].ToString();
                cname = drow2["c_name"].ToString();
                Worksheet sheet1 = workbook.Worksheets.Add(cname);
                exporttoexcel(page, conn, sheet1, style1, uniqueid, exportype);
            
            }
            workbook.Worksheets.RemoveAt(0);

            System.IO.MemoryStream ms = workbook.SaveToStream();//生成数据流
            byte[] bt = ms.ToArray();
            string filepath = page.Server.MapPath("/") + @"upload\attach\" + filename + ".xls";
            if (File.Exists(filepath)) System.IO.File.Delete(filepath);
            workbook.Save(filepath);//保存到硬盘

            success = @"\upload\attach\" + filename + ".xls";
            //page.Response.Write(filename);
        }

        private void exporttoexcel(HttpContext page, SqlConnection conn, Worksheet sheet, Style style1, string uniqueid, string exportype)
        {
            string year1 = "";
            string month1 = "";
            string region1 = "";

            if (page.Request["year1"] != null)
            {
                year1 = Functions.tostr(page.Request["year1"].ToString()).Trim();
            }

            if (page.Request["month1"] != null)
            {
                month1 = Functions.tostr(page.Request["month1"].ToString()).Trim();
            }

            if (page.Request["region1"] != null)
            {
                region1 = Functions.tostr(page.Request["region1"].ToString()).Trim();
            }


            Cells cells = sheet.Cells;
            string sql = "";
            sql = "exec qry_export_metricsdata '" + uniqueid + "'," + exportype + ",'" + year1 + "','" + month1 + "','" + region1 + "'";
            DataSet ds= loadfiexddata( conn,  sql);

            int row = 0;
            int column = 0;

            int column1 = ds.Tables[0].Columns.Count ;
            string[] arr = new string[column1];

            foreach (DataColumn item in ds.Tables[0].Columns)
            {
                arr[column] = item.ColumnName;
                column = column + 1;
            }
            column = 0;

            string value = "";

            foreach (DataRow drow1 in ds.Tables[0].Rows)
            {
                column = 0;
                style1.Custom = "@";
                if (drow1["tid"].ToString() == uniqueid)
                {
                    cells[row, column++].PutValue("序号");
                }
                else
                {
                    cells[row, column++].PutValue(row);
                }

                for (int i = 1; i < column1; i++)
                {
                    value = drow1[arr[i]].ToString();
                    if (row != 0)
                    {
                        if (arr[i] == "tyear" || arr[i] == "tmonth")
                        {
                            cells[row, column++].PutValue(int.Parse(value));
                        }
                        else
                        {
                            if (IsNumeric(value))
                                cells[row, column++].PutValue(double.Parse(value));
                            else cells[row, column++].PutValue(value);
                        }
                    }
                    else
                    {
                        cells[row, column++].PutValue(value);
                    }
                   
                    
                    
                }

                row = row + 1;
            }

        }

        
        private bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }

        private DataSet loadfiexddata(SqlConnection conn, string sql)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = sql;
            cmd.Connection = conn;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            da.Dispose();
            cmd.Dispose();
            return ds;
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