using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.NoFrame
{
    public class NoFrameController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public NoFrameController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "SaveDocument()", 1);
            pageofficeCtrl.AddCustomToolButton("打印设置", "PrintSet()", 0);
            pageofficeCtrl.AddCustomToolButton("打印", "PrintFile()", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen()", 4);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("关闭", "Close()", 21);

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
            fs.SaveToFile(webRootPath + "/NoFrame/doc/" + fs.FileName);
            return  fs.Close();
            
        }
    }
}