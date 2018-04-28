using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;

namespace DataOpt
{
    public class ToolsFunction
    {

        #region 字符串与对象的相互转换
        /// <summary>
        /// dataset 转为 json字符串
        /// </summary>
        /// <param name="dsData">dataset数据集</param>
        /// <returns>返回json数据</returns>
        public static string dataSetToJson(DataSet dsData)
        {
            if (dsData == null || dsData.Tables.Count <= 0 || dsData.Tables[0].Rows.Count <= 0)
            {
                //如果查询到的数据为空则返回标记ok:false
                return "{\"ok\":false}";
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("{\"ok\":true,");
            foreach (DataTable dt in dsData.Tables)
            {
                sb.Append(string.Format("\"{0}\":[", dt.TableName));

                foreach (DataRow dr in dt.Rows)
                {
                    sb.Append("{");
                    for (int i = 0; i < dr.Table.Columns.Count; i++)
                    {
                        sb.AppendFormat("\"{0}\":\"{1}\",", dr.Table.Columns[i].ColumnName.Replace("\"", "\\\"").Replace("\'", "\\\'"), ObjToStr(dr[i]).Replace("\"", "\\\"").Replace("\'", "\\\'")).Replace(Convert.ToString((char)13), "\\r\\n").Replace(Convert.ToString((char)10), "\\r\\n");
                    }
                    sb.Remove(sb.ToString().LastIndexOf(','), 1);
                    sb.Append("},");
                }

                sb.Remove(sb.ToString().LastIndexOf(','), 1);
                sb.Append("],");
            }
            sb.Remove(sb.ToString().LastIndexOf(','), 1);
            sb.Append("}");
            return sb.ToString();
        }

        /// <summary>
        /// 对象转字符串
        /// </summary>
        /// <param name="ob">对象</param>
        /// <returns>返回字符串</returns>
        public static string ObjToStr(object ob)
        {
            if (ob == null)
            {
                return string.Empty;
            }
            else
                return ob.ToString();
        }

        #endregion


    }
}
