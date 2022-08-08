using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ReadOnly
{
    public class ReadOnlyController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.Caption = "演示：文件在线安全浏览";
            pageofficeCtrl.JsFunction_AfterDocumentOpened = "AfterDocumentOpened()";
            pageofficeCtrl.AllowCopy = false;//禁止拷贝
            pageofficeCtrl.Menubar = false;
            pageofficeCtrl.OfficeToolbars = false;
            pageofficeCtrl.CustomToolbar = false;

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }
    }
}