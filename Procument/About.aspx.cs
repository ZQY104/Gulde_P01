using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

//using Excel = Microsoft.Office.Interop.Excel; //由于需要导出EXCEL, 需要增加此名称空间;
using System.Text;

namespace Procument
{
    public partial class About : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            ///<summary>
            ///思路
            ///1. 连接数据库
            ///2. 根据BOM number 查找出其对应的EMstring, 并放入一个Dataset中; 此时dataset.table[0]就是存放BOM number和其EM string 的表了.
            ///3. 设计一个FOR循环, 依次读出dataset.table[0]每一列的EMstring, 并执行EM string查询,从而获得partnumber, 将Part number存入另一个dataset中.
            ///3.1 在第三步中, 同时要增加一行用来显示BOM number;
            ///4. 设计一个按钮,用来输出gridview到EXCEL中.
            /// </summary>

            //数据库连接语句;
            string SQLString = ConfigurationManager.ConnectionStrings["GuldeERP_TestConnectionString"].ToString();//数据库连接语句设定为SQLString，该语句连接的是名为GuldeERP_Test的数据库；
            SqlConnection conn = new SqlConnection(SQLString);//建立一个名为conn的新连接，其语句为SQLString;
            conn.Open();//打开连接；
            //

            string Bomnumber = "5400-4";// "5400 -8WAT1E4P33T1111KKK5400-8WBT1E4P33T1111KKK5400-8WCT1E4P33T1111KKK5400-8WAT2E4P33T1111";
            string sql = "SELECT * FROM view_40bomquote where [Bom #] like ('" + Bomnumber + "%')";
            SqlCommand cmd = new SqlCommand(sql, conn);//建立一个SQL查询命令;
            SqlDataAdapter da = new SqlDataAdapter();//建立一个dataadapter;
            DataSet ds = new DataSet();//建立一个dataset;
            da.SelectCommand = cmd;//执行sql查询命令,并将查询结果注入到dataadapter;
            da.Fill(ds);//将查询结果注入到dataset中;

            #region GridView2查询

            string emstring = "";
            DataSet ds1 = new DataSet();//建立一个dataset;
            da.Fill(ds1); //先给ds1装入一个table[0], 如果没有这一行,则出现"cannot find table 0"错误提示;
            for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
            {
                //da.Fill(ds1);
                emstring = Convert.ToString((ds.Tables[0]).Rows[n][1].ToString());//将ds的第n行第2列赋值给emstring;

                DataRow dr = ds1.Tables[0].NewRow();//增加datarow用来显示BOM number
                dr[0] = Convert.ToString((ds.Tables[0]).Rows[n][0].ToString());
                ds1.Tables[0].Rows.Add(dr);

                string Emstring = emstring;
                //string Emstring = "5400X15-AG1-B20-C526-E17-F26-G4-H65-9A7-9D17-9E40";

                string[] Empick = Regex.Split(Emstring, "-", RegexOptions.IgnoreCase);//字符串Emstring被打散;

                string endstr = "";//新建字符串endstr,该字符串将包含除Emstring的第一个字符串外的所有字符串，例如Emstring为"5400X15-AG1-B42-C526",则endstr为"AG1,B42,C526";
                for (int i = 1; i < Empick.Length; i++)
                {
                    string str = "'" + Empick[i] + "'";  //在每个元素前后加上我们想要的格式，效果例如：  // " 'AG1' "
                    if (i < Empick.Length - 1)  //根据数组元素的个数来判断应该加多少个逗号
                    {
                        str += ",";
                    }
                    endstr += str;//将str累加的字符串赋值给endstr;这样就得到字符串为endstr = "'AG1','B42','C526'";
                }

                //查询em string并注入到GridView中;
                string sql1 = "Select * from View_40structure where (Em = '" + Empick[0] + "' and Empick in (" + endstr + ") )order by Empick";//指定范围查询并排序;
                SqlCommand cmd1 = new SqlCommand(sql1, conn);//建立一个SQL查询命令;
                SqlDataAdapter da1 = new SqlDataAdapter();  //建立一个dataadapter;
                da1.SelectCommand = cmd1;//执行sql查询命令,并将查询结果注入到dataadapter;
                da1.Fill(ds1);//将查询结果注入到dataset中;
            }

            this.GridView1.DataSource = ds1;//将dataset结果注入到gridview;
            this.GridView1.DataBind();//对gridview进行数据绑定;

            #endregion GridView2查询
        }

        /// <summary>
        /// 输出GridView1到excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Button2_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now; //获取server当前时间;
            Export("application/ms-excel", "BOM List" + dt.ToString() + ".xls");//将导出的文件名命名为BOM LIST(当前时间).xls;
        }

        private void Export(string FileType, string FileName)
        {
            Response.Charset = "GB2312";//避免文件名为乱码
            Response.ContentEncoding = System.Text.Encoding.UTF8; //不知道干啥的，修改这行UTF8到7就出现乱码，最好别加了。
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(FileName, Encoding.UTF8).ToString());
            Response.ContentType = FileType;
            this.EnableViewState = false;
            StringWriter tw = new StringWriter(); //实例化StringWriter;
            HtmlTextWriter hw = new HtmlTextWriter(tw); //实例化HtmlTextWriter;
            //GridView1.Columns[7].Visible = false; //**!!!!**将第7列，也就是按钮列隐藏,否则无法导出来;
            GridView1.RenderControl(hw);
            //GridView1.RenderControl(hw);
            Response.Write(tw.ToString());
            Response.End();
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
        }
    }
}