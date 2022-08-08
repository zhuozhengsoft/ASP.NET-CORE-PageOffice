using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordCompare
{
    public class WordCompareController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            // Create custom toolbar
            pageofficeCtrl.AddCustomToolButton("保存", "SaveDocument()", 1);
            pageofficeCtrl.AddCustomToolButton("显示A文档", "ShowFile1View()", 0);
            pageofficeCtrl.AddCustomToolButton("显示B文档", "ShowFile2View()", 0);
            pageofficeCtrl.AddCustomToolButton("显示比较结果", "ShowCompareView()", 0);


            pageofficeCtrl.WordCompare("doc/aaa1.doc", "doc/aaa2.doc", PageOfficeNetCore.OpenModeType.docAdmin, "Tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}