using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.OpenWord
{
    public class OpenWordController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //隐藏Office工具条
            pageofficeCtrl.OfficeToolbars = false;
            //隐藏自定义工具栏
            pageofficeCtrl.CustomToolbar = false;
            //设置页面的显示标题
            pageofficeCtrl.Caption = "演示：最简单的以只读模式打开Word文档";

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}