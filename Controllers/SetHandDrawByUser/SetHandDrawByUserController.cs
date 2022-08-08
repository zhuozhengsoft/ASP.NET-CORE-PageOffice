using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SetHandDrawByUser
{
    public class SetHandDrawByUserController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        private readonly IWebHostEnvironment _webHostEnvironment;
        public SetHandDrawByUserController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            string userName = "";
            userName = Request.Form["userName"];
            if ("zhangsan" == userName) userName = "张三";
            if ("lisi" == userName) userName = "李四";
            if ("wangwu" == userName) userName = "王五";
            //***************************卓正PageOffice组件的使用********************************

            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("领导圈阅", "StartHandDraw", 3);
            //pageofficeCtrl.AddCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.JsFunction_AfterDocumentOpened = "ShowByUserName";

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.userName = userName;
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SetHandDrawByUser/doc/" + fs.FileName);
            return fs.Close();
            
        }

    }
}