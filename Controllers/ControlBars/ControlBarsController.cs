using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ControlBars
{
    public class ControlBarsController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.Titlebar = false;
            pageofficeCtrl.Menubar = false;
            pageofficeCtrl.OfficeToolbars = false;
            pageofficeCtrl.CustomToolbar = false;
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}