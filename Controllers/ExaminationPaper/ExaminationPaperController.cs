using System;
using System.Data.SQLite;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;

using Microsoft.AspNetCore.Mvc;


namespace NetCoreSamples5.Controllers.ExaminationPaper
{
    public class ExaminationPaperController : Controller
    {

        private string connString;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public ExaminationPaperController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath= _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "ExaminationPaper.db";
            connString = "Data Source=" + dataPath;
        }


        public IActionResult Index()
        {
            string sql = "Select * from stream";
            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SQLiteDataReader dr = cmd.ExecuteReader();

            StringBuilder strHtmls = new StringBuilder();

            strHtmls.Append("<tr  style='background-color:#FEE;'>");
            strHtmls.Append("<td style='text-align:center;width=10%' >选择</td>");
            strHtmls.Append("<td style='text-align:center;width=30%'>题库编号</td>");
            strHtmls.Append("<td style='text-align:center;width=60%'>操作</td>");
            strHtmls.Append("</tr>");

            if (dr.Read())
            {
                string pID = dr["ID"].ToString().Trim();
                strHtmls.Append("<tr  style='background-color:white;'>");
                strHtmls.Append("<td><input id='check" + pID + "'  type='checkbox' /></td>");
                strHtmls.Append("<td>选择题-" + pID + "</td>");
                strHtmls.Append("<td><a href='javascript:POBrowser.openWindowModeless(\"Edit?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
                strHtmls.Append("</tr>");

                while (dr.Read())
                {
                    pID = dr["ID"].ToString().Trim();
                    strHtmls.Append("<tr  style='background-color:white;'>");
                    strHtmls.Append("<td><input id='check" + pID + "'  type='checkbox' /></td>");
                    strHtmls.Append("<td>选择题-" + pID + "</td>");
                    strHtmls.Append("<td><a href='javascript:POBrowser.openWindowModeless(\"Edit?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
                    strHtmls.Append("</tr>");
                }
            }
            else
            {
                strHtmls.Append("<tr>\r\n");
                strHtmls.Append("<td colspan='3' width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
                strHtmls.Append("</td></tr>\r\n");
            }

            ViewBag.strHtmls = strHtmls.ToString();
            dr.Close();
            conn.Close();
            return View();
        }

        public IActionResult Edit()
        {

            string id = Request.Query["id"];
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            pageofficeCtrl.AddCustomToolButton("保存","Save",1);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc?id=" + id;
            //打开Word文档
            pageofficeCtrl.WebOpen("Openfile?id=" + id, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult Compose()
        {

            string operateStr = "";

            string idlist = Request.Query["ids"];
            string[] ids = idlist.Split(',');//将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件即可
            int pNum = 1;

            operateStr +="function Create(){\n";
            // document.getElementById('PageOfficeCtrl1').Document.Application 微软office VBA对象的根Application对象
            operateStr += "var obj = document.getElementById('PageOfficeCtrl1').Document.Application;\n";
            operateStr += "obj.Selection.EndKey(6);\n"; // 定位光标到文档末尾

            for (int i = 0; i < ids.Length; i++)
            {

                operateStr += "obj.Selection.TypeParagraph();"; //用来换行
                operateStr += "obj.Selection.Range.Text = '" + pNum.ToString() + ".';\n"; // 用来生成题号
                // 下面两句代码用来移动光标位置
                operateStr += "obj.Selection.EndKey(5,1);\n";
                operateStr += "obj.Selection.MoveRight(1,1);\n";
                // 插入指定的题到文档中
                operateStr += "document.getElementById('PageOfficeCtrl1').InsertDocumentFromURL('Openfile?id=" + ids[i] + "');\n";
                pNum++;

            }
            operateStr += "\n}\n";

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            pageofficeCtrl.Caption = "生成试卷";
            pageofficeCtrl.JsFunction_AfterDocumentOpened = "Create";
            pageofficeCtrl.CustomToolbar = false;
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.operateStr = operateStr;
            return View();
        }

        public IActionResult Compose2()
        {
            int num = 1;//试题编号

            string idlist = Request.Query["ids"];
            string[] ids = idlist.Split(',');//将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件即可

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            string temp = "PO_begin";//新添加的数据区域名称

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            for (int i = 0; i < ids.Length; i++)
            {

                PageOfficeNetCore.WordWriter.DataRegion dataNum = doc.CreateDataRegion("PO_" + num, PageOfficeNetCore.WordWriter.DataRegionInsertType.After, temp);
                dataNum.Value = num + ".\t";
                PageOfficeNetCore.WordWriter.DataRegion dataReg = doc.CreateDataRegion("PO_begin" + (i + 1), PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "PO_" + num);
                dataReg.Value = "[word]Openfile?id=" + ids[i] + "[/word]";
                num++;
                temp = "PO_begin" + (i + 1);
            }
            pageofficeCtrl.SetWriter(doc);
            pageofficeCtrl.Caption = "生成试卷";
            pageofficeCtrl.CustomToolbar = false;
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public void Openfile()
        {
            string docID = Request.Query["id"];
            string sql = "select Word from stream where id =" + docID;
            using (SQLiteConnection conn = new SQLiteConnection(connString))
            {
                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                    SQLiteDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        long num = dr.GetBytes(0, 0, null, 0, Int32.MaxValue);
                        Byte[] b = new Byte[num];
                        dr.GetBytes(0, 0, b, 0, b.Length);
                        Response.ContentType = "Application/msword"; //其他文件格式换成相应类型即可 application/x-excel, application/ms-powerpoint, application/pdf 
                        Response.Headers.Add("Content-Disposition", "attachment; filename=down.doc");//其他文件格式换成相应类型的filename
                        Response.Headers.Add("Content-Length", num.ToString());
                        Response.Body.WriteAsync(b);
                    }
                }

            }
            Response.Body.Flush();
            Response.Body.Close();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string id = Request.Query["id"];
            string sql = "UPDATE  Stream SET Word=@file  where ID=" + id;
            using (SQLiteConnection conn = new SQLiteConnection(connString))
            {
                conn.Open();
                byte[] aa = fs.FileBytes;
                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                {
                    SQLiteParameter parameter = new SQLiteParameter("@file",System.Data.DbType.Binary);
                    parameter.Value = aa;
                    cmd.Parameters.Add(parameter);
                    cmd.ExecuteNonQuery();
                }
            }
            return fs.Close();
            
        }

    }
}