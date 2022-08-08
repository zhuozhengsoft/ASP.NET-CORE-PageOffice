using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.Data.SQLite;

namespace NetCoreSamples5.Controllers.DataBase
{
    public class DataBaseController : Controller
    {

        private String connString ;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public DataBaseController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "DataBase.db";
            connString = "Data Source="+ dataPath ;

        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc?id=1";
            //打开Word文档
            pageofficeCtrl.WebOpen("Openstream?id=1", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
        public void Openstream()
        {
            string docID = Request.Query["id"];
            string sql = "Select Word,ID  from stream where ID='" + docID + "'";
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

            string sID = Request.Query["id"];

            string sql = "UPDATE  Stream SET Word=@file WHERE ID=" + sID;

            using (SQLiteConnection conn = new SQLiteConnection(connString))
            {
                conn.Open();
                byte[] aa = fs.FileBytes;

                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                {
                    SQLiteParameter parameter = new SQLiteParameter("@file", System.Data.DbType.Binary);
                    parameter.Value = aa;
                    cmd.Parameters.Add(parameter);
                    cmd.ExecuteNonQuery();
                }

            }
            return fs.Close();
            
        }

    }
}