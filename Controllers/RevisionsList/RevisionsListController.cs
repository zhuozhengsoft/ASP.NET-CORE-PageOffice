using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.RevisionsList
{
    public class RevisionsListController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public RevisionsListController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.JsFunction_AfterDocumentOpened = "AfterDocumentOpened()";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.OfficeToolbars = false;

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docRevisionOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/RevisionsList/doc/" + fs.FileName);
            return  fs.Close();
            
        }
    }
}