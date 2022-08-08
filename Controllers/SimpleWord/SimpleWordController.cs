using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SimpleWord
{
    public class SimpleWordController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public SimpleWordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("另存到本地", "SaveAs", 12);
            pageofficeCtrl.AddCustomToolButton("打印设置", "PrintSet", 0);
            pageofficeCtrl.AddCustomToolButton("打印", "PrintFile", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("关闭", "Close", 21);
            pageofficeCtrl.FileTitle="test";//设置另存为到本地的文件名称
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult Word1()

        {

            string simplePath = "\\SimpleWord\\doc\\test.doc";
            string finalPath = _webHostEnvironment.WebRootPath + simplePath;

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen(finalPath, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SimpleWord/doc/" + fs.FileName);
            return   fs.Close();
            
        }
    }
}