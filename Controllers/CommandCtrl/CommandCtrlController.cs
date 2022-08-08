using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.CommandCtrl
{
    public class CommandCtrlController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.CustomToolbar = false;
            pageofficeCtrl.OfficeToolbars = false;
            // 禁止拷贝
            pageofficeCtrl.AllowCopy = false;

            pageofficeCtrl.JsFunction_AfterDocumentOpened = "AfterDocumentOpened";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}