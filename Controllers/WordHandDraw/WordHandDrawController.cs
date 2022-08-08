using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordHandDraw
{
    public class WordHandDrawController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public WordHandDrawController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("开始手写", "StartHandDraw()", 5);
            pageofficeCtrl.AddCustomToolButton("设置线宽", "SetPenWidth()", 5);
            pageofficeCtrl.AddCustomToolButton("设置颜色", "SetPenColor()", 5);
            pageofficeCtrl.AddCustomToolButton("设置笔型", "SetPenType()", 5);
            pageofficeCtrl.AddCustomToolButton("设置缩放", "SetPenZoom()", 5);

            pageofficeCtrl.AddCustomToolButton("访问手写集", "GetHandDrawList()", 6);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docRevisionOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/WordHandDraw/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}