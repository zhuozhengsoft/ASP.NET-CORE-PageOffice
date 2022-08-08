using System;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.CreateWord
{
    public class CreateWordController : Controller
    {
        private string connString;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public CreateWordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "CreateWord.db";
            connString = "Data Source=" + dataPath;
        }

        public IActionResult Index()
        {

            string op = Request.Query["op"];
            string FileSubject = Request.Query["FileSubject"];
            if (op != null && op.Length > 0)
            {
                Insert(FileSubject);
            }

            string sql = "select * from word order by ID desc ";
            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SQLiteDataReader dr = cmd.ExecuteReader();
            StringBuilder strHtml = new StringBuilder();
            while (dr.Read())
            {
                strHtml.Append("<tr onmouseover='onColor(this)' onmouseout='offColor(this)'>\n");
                strHtml.Append("<td><a href =\"javascript:POBrowser.openWindowModeless('Word?filename="
                    + dr["FileName"].ToString() + "&subject="
                    + dr["Subject"].ToString() + "','width=1200px;height=800px;');\">"
                    + dr["Subject"].ToString() + "</a></td>\n");
                if (dr["SubmitTime"].ToString() != "" )
                {
                    strHtml.Append("<td>" + DateTime.Parse(dr["SubmitTime"].ToString()).ToString("yyyy-MM-dd") + "</td>\n");
                }
                else
                {
                    strHtml.Append("<td>&nbsp;</td>\n");
                }
                strHtml.Append(" </tr>\n");
            }
            ViewBag.strHtml = strHtml.ToString();
            return View();
        }

        private String  Insert(string FileSubject)
        {

            string newID = "";
            string sql = "select Max(ID) from word";
            SQLiteConnection conn = new SQLiteConnection(connString);
            conn.Open();
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SQLiteDataReader dr = cmd.ExecuteReader();

            if (dr.Read() && dr[0].ToString().Trim().Length > 0)
            {
                newID = (Convert.ToInt32(dr[0]) + 1).ToString();
            }
            dr.Close();
            string fileName = "aabb" + newID + ".doc";

            string strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
            + ",'" + fileName + "','" + FileSubject + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "')";
            cmd.CommandText = strsql;
            cmd.ExecuteNonQuery();
            conn.Close();

            return fileName;
        }

        public IActionResult Word()
        {

            string fileName = Request.Query["filename"];
            string subject = Request.Query["subject"];

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/" + fileName, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.subject = subject;
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult CreateWord()
        {
            string fileName = Request.Query["fileName"];
            string subject = Request.Query["subject"];
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.JsFunction_BeforeDocumentSaved = "BeforeDocumentSaved()";
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveNewDoc";
            //打开Word文档
            pageofficeCtrl.WebCreateNew("张佚名", PageOfficeNetCore.DocumentVersion.Word2003);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/CreateWord/doc/" + fs.FileName);
            fs.CustomSaveResult = "ok";
            return fs.Close();
            
        }

        public async Task<ActionResult> SaveNewDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string subject = fs.GetFormField("FileSubject");
             string fileName=Insert(subject);//向数据库插入文件记录并返回文件名称
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/CreateWord/doc/" + fileName);
            fs.CustomSaveResult = "ok";
            return fs.Close();
            
        }
    }
}