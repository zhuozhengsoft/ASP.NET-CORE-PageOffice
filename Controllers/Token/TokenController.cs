
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.Token
{
    public class TokenController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public TokenController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Word()
        {
            //获取token值
            string mytoken = Request.Headers["Authorization"];
            ViewBag.testToken = mytoken;

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.AddCustomToolButton("保存","Save",1);
            pageofficeCtrl.ServerPage = "/POserver";
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            string mytoken = Request.Headers["Authorization"];
    
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/Token/doc/" + fs.FileName);
            fs.ShowPage(400, 300, this);
            fs.Close();

            ViewBag.testToken = mytoken;
            return View();
        }
    }
}