using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.BeforeAndAfterSave
{
    public class BeforeAndAfterSaveController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public BeforeAndAfterSaveController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            // 设置文件保存之前执行的事件
            pageofficeCtrl.JsFunction_BeforeDocumentSaved = "BeforeDocumentSaved()";
            // 设置文件保存之后执行的事件
            pageofficeCtrl.JsFunction_AfterDocumentSaved = "AfterDocumentSaved()";

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/BeforeAndAfterSave/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}