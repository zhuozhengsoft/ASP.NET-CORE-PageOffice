using System;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordSalaryBill
{


    public class WordSalaryBillController : Controller
    {

        private String connString;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public WordSalaryBillController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "WordSalaryBill.db";
            connString = "Data Source=" + dataPath;

        }


        public IActionResult Index()
        {

            string sql = "select * from Salary   order by ID ";
            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SQLiteDataReader dr = cmd.ExecuteReader();
            StringBuilder strHtmls = new StringBuilder();
            strHtmls.Append("<tr  style='height:40px;padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; background:#edf8fe; font-weight:bold; color:#666;text-align:center; text-indent:0px;'>");
            strHtmls.Append("<td width='5%' >选择</td>");
            strHtmls.Append("<td width='10%' >员工编号</td>");
            strHtmls.Append("<td width='10%' >员工姓名</td>");
            strHtmls.Append("<td width='15%' >所在部门</td>");
            strHtmls.Append("<td width='10%' >应发工资</td>");
            strHtmls.Append("<td width='10%' >扣除金额</td>");
            strHtmls.Append("<td width='10%' >实发工资</td>");
            strHtmls.Append("<td width='10%' >发放日期</td>");
            strHtmls.Append("<td width='20%' >操作</td>");
            strHtmls.Append("</tr>");

            bool flg = false;

            while (dr.Read())
            {
                flg = true;
                DateTime date = DateTime.Now;
                string pID = dr["ID"].ToString().Trim();
                strHtmls.Append("<tr  style='height:40px; text-indent:10px; padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; color:#666;'>");
                strHtmls.Append("<td style=' text-align:center;'><input id='check" + pID + "'  type='checkbox' /></td>");
                strHtmls.Append("<td style=' text-align:left;'>" + pID + "</td>");
                strHtmls.Append("<td style=' text-align:left;'>" + dr["UserName"].ToString() + "</td>");
                strHtmls.Append("<td style=' text-align:left;'>" + dr["DeptName"].ToString() + "</td>");
                if (dr["SalTotal"] != null && dr["SalTotal"].ToString() != "")
                {
                    strHtmls.Append("<td style=' text-align:left;'>" + dr["SalTotal"].ToString() + "</td>");
                }
                else
                {
                    strHtmls.Append("<td style=' text-align:left;'>￥0.00</td>");
                }

                if (dr["SalDeduct"] != null && dr["SalDeduct"].ToString() != "")
                {
                    strHtmls.Append("<td style=' text-align:left;'>" + dr["SalDeduct"].ToString() + "</td>");
                }
                else
                {
                    strHtmls.Append("<td style=' text-align:left;'>￥0.00</td>");
                }

                if (dr["SalCount"] != null && dr["SalCount"].ToString() != "")
                {
                    strHtmls.Append("<td style=' text-align:left;'>" + dr["SalCount"].ToString() + "</td>");
                }
                else
                {
                    strHtmls.Append("<td style=' text-align:left;'>￥0.00</td>");
                }

                if (dr["DataTime"] != null && dr["DataTime"].ToString() != "")
                {
                    strHtmls.Append("<td style=' text-align:center;'>" + dr["DataTime"].ToString() + "</td>");
                }
                else
                {
                    strHtmls.Append("<td style=' text-align:left;'>" + DateTime.Now.ToString("yyyy-MM-dd") + "</td>");
                }
                strHtmls.Append("<td style=' text-align:center;'><a href='javascript:POBrowser.openWindowModeless(\"ViewWord?ID=" + pID + "\" ,\"width=1200px;height=800px;\");' >查看</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:POBrowser.openWindowModeless(\"Openfile?ID=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
                strHtmls.Append("</tr>");
            }
            if (!flg)
            {
                strHtmls.Append("<tr>\r\n");
                strHtmls.Append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
                strHtmls.Append("</td></tr>\r\n");
            }
            ViewBag.strHtmls = strHtmls.ToString();
            dr.Close();
            conn.Close();
            return View();
        }

        public IActionResult ViewWord()
        {
            String err = "";
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            string id = Request.Query["ID"];

            if (id != null && id.Length > 0)
            {
                string sql = "select * from Salary where id =" + id + " order by ID"; ;
                SQLiteConnection conn = new SQLiteConnection(connString);
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                cmd.ExecuteNonQuery();
                cmd.CommandText = sql;
                SQLiteDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    DateTime date = DateTime.Now;

                    //创建WordDocment对象
                    PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
                    //打开数据区域
                    PageOfficeNetCore.WordWriter.DataRegion datareg = doc.OpenDataRegion("PO_table");
                    //打开Table
                    PageOfficeNetCore.WordWriter.Table table = datareg.OpenTable(1);
                    ////给单元格赋值
                    table.OpenCellRC(2, 1).Value = dr["ID"].ToString();
                    table.OpenCellRC(2, 2).Value = dr["UserName"].ToString();
                    table.OpenCellRC(2, 3).Value = dr["DeptName"].ToString();

                    if (dr["SalTotal"] != null && dr["SalTotal"].ToString() != "")
                    {
                        table.OpenCellRC(2, 4).Value = dr["SalTotal"].ToString();
                    }
                    else
                    {
                        table.OpenCellRC(2, 4).Value = "￥0.00";
                    }

                    if (dr["SalDeduct"] != null && dr["SalDeduct"].ToString() != "")
                    {
                        table.OpenCellRC(2, 5).Value = dr["SalDeduct"].ToString();
                    }
                    else
                    {
                        table.OpenCellRC(2, 5).Value = "￥0.00";
                    }

                    if (dr["SalCount"] != null && dr["SalCount"].ToString() != "")
                    {
                        table.OpenCellRC(2, 6).Value = dr["SalCount"].ToString();
                    }
                    else
                    {
                        table.OpenCellRC(2, 6).Value = "￥0.00";
                    }

                    if (dr["DataTime"] != null && dr["SalTotal"].ToString() != "")
                    {
                        table.OpenCellRC(2, 7).Value = dr["DataTime"].ToString();
                    }
                    else
                    {
                        table.OpenCellRC(2, 7).Value = "";
                    }

                    pageofficeCtrl.SetWriter(doc);
                }
                else
                {
                    err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
                }
                dr.Close();
                conn.Close();

                //打开Word文档

            }
            else
            {
                err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
            }
            pageofficeCtrl.CustomToolbar = false;
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.err = err;
            return View();
        }



        public IActionResult Compose()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            string idlist = Request.Query["ids"];
            string sql = "select * from Salary where ID in(" + idlist + ") order by ID";

            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SQLiteDataReader dr = cmd.ExecuteReader();

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            PageOfficeNetCore.WordWriter.DataRegion data = null;
            PageOfficeNetCore.WordWriter.Table table = null;
            int i = 0;
            while (dr.Read())
            {
                data = doc.CreateDataRegion("reg" + i.ToString(), PageOfficeNetCore.WordWriter.DataRegionInsertType.Before, "[End]");
                data.Value = "[word]doc/template.doc[/word]";
                table = data.OpenTable(1);

                table.OpenCellRC(2, 1).Value = dr["ID"].ToString();
                table.OpenCellRC(2, 2).Value = dr["UserName"].ToString();
                table.OpenCellRC(2, 3).Value = dr["DeptName"].ToString();

                if (dr["SalTotal"] != null && dr["SalTotal"].ToString() != "")
                {
                    table.OpenCellRC(2, 4).Value = dr["SalTotal"].ToString();
                }
                else
                {
                    table.OpenCellRC(2, 4).Value = "￥0.00";
                }

                if (dr["SalDeduct"] != null && dr["SalDeduct"].ToString() != "")
                {
                    table.OpenCellRC(2, 5).Value = dr["SalDeduct"].ToString();
                }
                else
                {
                    table.OpenCellRC(2, 5).Value = "￥0.00";
                }

                if (dr["SalCount"] != null && dr["SalCount"].ToString() != "")
                {
                    table.OpenCellRC(2, 6).Value = dr["SalCount"].ToString();
                }
                else
                {
                    table.OpenCellRC(2, 6).Value = "￥0.00";
                }

                if (dr["DataTime"] != null && dr["SalTotal"].ToString() != "")
                {
                    table.OpenCellRC(2, 7).Value = dr["DataTime"].ToString();
                }
                else
                {
                    table.OpenCellRC(2, 7).Value = "";
                }
                i++;
            }

            dr.Close();
            conn.Close();

            // 设置PageOffice组件服务页面
            pageofficeCtrl.SetWriter(doc);

            pageofficeCtrl.Caption = "生成工资条";
            pageofficeCtrl.CustomToolbar = false;
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docAdmin, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }



        public IActionResult Openfile()
        {
            String err = "";
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            string id = Request.Query["ID"];

            if (id != null && id.Length > 0)
            {
                string sql = "select * from Salary where id =" + id + " order by ID"; ;
                SQLiteConnection conn = new SQLiteConnection(connString);
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(sql, conn);
                cmd.ExecuteNonQuery();
                cmd.CommandText = sql;
                SQLiteDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    DateTime date = DateTime.Now;

                    //创建WordDocment对象
                    PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
                    //打开数据区域
                    PageOfficeNetCore.WordWriter.DataRegion datareg = doc.OpenDataRegion("PO_table");
                    //给数据区域赋值
                    doc.OpenDataRegion("PO_ID").Value = id;

                    //设置数据区域的可编辑性
                    doc.OpenDataRegion("PO_UserName").Editing = true;
                    doc.OpenDataRegion("PO_DeptName").Editing = true;
                    doc.OpenDataRegion("PO_SalTotal").Editing = true;
                    doc.OpenDataRegion("PO_SalDeduct").Editing = true;
                    doc.OpenDataRegion("PO_SalCount").Editing = true;
                    doc.OpenDataRegion("PO_DataTime").Editing = true;

                    doc.OpenDataRegion("PO_UserName").Value = dr["UserName"].ToString();
                    doc.OpenDataRegion("PO_DeptName").Value = dr["DeptName"].ToString();


                    if (dr["SalTotal"] != null && dr["SalTotal"].ToString() != "")
                    {
                        doc.OpenDataRegion("PO_SalTotal").Value = dr["SalTotal"].ToString();
                    }
                    else
                    {
                        doc.OpenDataRegion("PO_SalTotal").Value = "￥0.00";
                    }

                    if (dr["SalDeduct"] != null && dr["SalDeduct"].ToString() != "")
                    {
                        doc.OpenDataRegion("PO_SalDeduct").Value = dr["SalDeduct"].ToString();
                    }
                    else
                    {
                        doc.OpenDataRegion("PO_SalDeduct").Value = "￥0.00";
                    }

                    if (dr["SalCount"] != null && dr["SalCount"].ToString() != "")
                    {
                        doc.OpenDataRegion("PO_SalCount").Value = dr["SalCount"].ToString();

                    }
                    else
                    {
                        doc.OpenDataRegion("PO_SalCount").Value = "￥0.00";
                    }

                    if (dr["DataTime"] != null && dr["DataTime"].ToString() != "")
                    {
                        doc.OpenDataRegion("PO_DataTime").Value = dr["DataTime"].ToString(); ;
                    }
                    else
                    {
                        doc.OpenDataRegion("PO_DataTime").Value = DateTime.Now.ToString("yyyy-MM-dd");
                    }

                    pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
                    pageofficeCtrl.SaveDataPage = "SaveData?id=" + id;
                    pageofficeCtrl.SetWriter(doc);
                }
                else
                {
                    err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
                }
                dr.Close();
                conn.Close();

                //打开Word文档

            }
            else
            {
                err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
            }

            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.err = err;
            return View();
        }

        public async Task<ActionResult> SaveData()
        {

            string id = Request.Query["id"];

            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();

            string userName = "", deptName = "", salTotoal = "0", salDeduct = "0", salCount = "0", dateTime = "";
            //-----------  PageOffice 服务器端编程开始  -------------------//
            userName = doc.OpenDataRegion("PO_UserName").Value;
            deptName = doc.OpenDataRegion("PO_DeptName").Value;
            salTotoal = doc.OpenDataRegion("PO_SalTotal").Value;
            salDeduct = doc.OpenDataRegion("PO_SalDeduct").Value;
            salCount = doc.OpenDataRegion("PO_SalCount").Value;
            dateTime = doc.OpenDataRegion("PO_DataTime").Value;

            string sql = "UPDATE Salary SET UserName='" + userName
                + "',DeptName='" + deptName + "',SalTotal='" + salTotoal
                + "',SalDeduct='" + salDeduct + "',SalCount='" + salCount
                + "',DataTime='" + dateTime + "' WHERE ID=" + id;

            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);

            cmd.ExecuteNonQuery();
            conn.Close();
            return doc.Close();
            
        }

    }
}